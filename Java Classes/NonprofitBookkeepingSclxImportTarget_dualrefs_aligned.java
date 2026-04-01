package nonprofitbookkeeping.importer.sclx;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import nonprofitbookkeeping.core.Database;
import nonprofitbookkeeping.model.*;
import nonprofitbookkeeping.model.impex.BankStatementRecord;
import nonprofitbookkeeping.model.impex.BankingItemRecord;
import nonprofitbookkeeping.model.impex.BudgetRecord;
import nonprofitbookkeeping.model.supplemental.*;
import nonprofitbookkeeping.persistence.AccountRepository;
import nonprofitbookkeeping.persistence.DocumentRepository;
import nonprofitbookkeeping.persistence.PersonRepository;
import nonprofitbookkeeping.persistence.impex.BankStatementRecordRepository;
import nonprofitbookkeeping.persistence.impex.BankingItemRecordRepository;
import nonprofitbookkeeping.persistence.impex.BudgetRecordRepository;
import nonprofitbookkeeping.persistence.supplemental.TxnSupplementalLineMapper;
import nonprofitbookkeeping.persistence.supplemental.TxnSupplementalLineRecord;
import nonprofitbookkeeping.persistence.supplemental.TxnSupplementalLineRepository;
import nonprofitbookkeeping.service.FundAccountingService;
import nonprofitbookkeeping.service.scaledger.JournalLedgerPersistenceGateway;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.SQLException;
import java.time.ZoneOffset;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Concrete SCLX import target for NonprofitBookkeeping.
 */
public class NonprofitBookkeepingSclxImportTarget implements SclxImportTarget
{
    private static final ObjectMapper MAPPER = new ObjectMapper().findAndRegisterModules();

    private final AccountRepository accountRepository;
    private final PersonRepository personRepository;
    private final JournalLedgerPersistenceGateway journalGateway;
    private final TxnSupplementalLineRepository supplementalRepository;
    private final BudgetRecordRepository budgetRecordRepository;
    private final BankingItemRecordRepository bankingItemRecordRepository;
    private final BankStatementRecordRepository bankStatementRecordRepository;
    private final DocumentRepository documentRepository;
    private final FundAccountingService fundAccountingService;

    private final Map<String, Integer> importedTransactionIdsBySclxId = new LinkedHashMap<>();
    private final Map<Integer, Integer> importedTransactionIdsByLedgerRow = new LinkedHashMap<>();
    private final Map<String, Long> personDbIdBySclxPersonId = new LinkedHashMap<>();
    private final Map<String, String> personDisplayNameBySclxPersonId = new LinkedHashMap<>();
    private SclxImportOptions currentOptions = SclxImportOptions.defaults();

    public NonprofitBookkeepingSclxImportTarget()
    {
        this(
            new AccountRepository(),
            new PersonRepository(),
            new JournalLedgerPersistenceGateway(),
            new TxnSupplementalLineRepository(),
            new BudgetRecordRepository(),
            new BankingItemRecordRepository(),
            new BankStatementRecordRepository(),
            new DocumentRepository(),
            new FundAccountingService()
        );
    }

    public NonprofitBookkeepingSclxImportTarget(
        AccountRepository accountRepository,
        PersonRepository personRepository,
        JournalLedgerPersistenceGateway journalGateway,
        TxnSupplementalLineRepository supplementalRepository,
        BudgetRecordRepository budgetRecordRepository,
        BankingItemRecordRepository bankingItemRecordRepository,
        BankStatementRecordRepository bankStatementRecordRepository,
        DocumentRepository documentRepository,
        FundAccountingService fundAccountingService)
    {
        this.accountRepository = Objects.requireNonNull(accountRepository);
        this.personRepository = Objects.requireNonNull(personRepository);
        this.journalGateway = Objects.requireNonNull(journalGateway);
        this.supplementalRepository = Objects.requireNonNull(supplementalRepository);
        this.budgetRecordRepository = Objects.requireNonNull(budgetRecordRepository);
        this.bankingItemRecordRepository = Objects.requireNonNull(bankingItemRecordRepository);
        this.bankStatementRecordRepository = Objects.requireNonNull(bankStatementRecordRepository);
        this.documentRepository = Objects.requireNonNull(documentRepository);
        this.fundAccountingService = Objects.requireNonNull(fundAccountingService);
    }

    @Override
    public void beginImport(SclxDocument document, SclxImportOptions options)
    {
        this.currentOptions = options == null ? SclxImportOptions.defaults() : options;
        importedTransactionIdsBySclxId.clear();
        importedTransactionIdsByLedgerRow.clear();
        personDbIdBySclxPersonId.clear();
        personDisplayNameBySclxPersonId.clear();
    }

    @Override
    public void importOrganization(SclxDocument.Organization organization)
    {
        upsertDocumentJson("sclx.organization", organization);
    }

    @Override
    public void importReportingPeriod(SclxDocument.ReportingPeriod reportingPeriod)
    {
        upsertDocumentJson("sclx.reportingPeriod", reportingPeriod);
    }

    @Override
    public void importAccounts(List<SclxDocument.Account> accounts)
    {
        for (SclxDocument.Account source : accounts)
        {
            Account account = new Account();
            account.setAccountNumber(resolveAccountNumber(source.accountId(), source.Number()));
            account.setName(firstNonBlank(source.Name(), source.accountId(), source.Number()));
            account.setIncreaseSide(AccountSide.fromString(nullToEmpty(source.IncreaseSide())));
            account.setAccountType(AccountType.fromString(nullToEmpty(source.Type())));
            account.setParentAccountId(source.Parent());
            account.setCurrency("USD");
            account.setOpeningBalance(source.OpeningBalance() == null ? BigDecimal.ZERO : source.OpeningBalance());
            if (source.SupplementalKinds() != null)
            {
                List<SupplementalLineKind> kinds = source.SupplementalKinds().stream()
                    .map(this::parseSupplementalKind)
                    .filter(Objects::nonNull)
                    .collect(Collectors.toCollection(ArrayList::new));
                account.setSupplementalLineKinds(kinds);
            }
            try
            {
                accountRepository.upsert(account);
            }
            catch (SQLException ex)
            {
                throw new IllegalStateException("Failed to persist account " + account.getAccountNumber(), ex);
            }
        }
    }

    @Override
    public void importFunds(List<SclxDocument.Fund> funds)
    {
        for (SclxDocument.Fund source : funds)
        {
            if (source == null || source.name() == null || source.name().isBlank())
            {
                continue;
            }
            try
            {
                fundAccountingService.addFund(new Fund(source.name()));
            }
            catch (IllegalArgumentException ignored)
            {
            }
        }
        try
        {
            fundAccountingService.saveFunds(null);
        }
        catch (Exception ex)
        {
            throw new IllegalStateException("Failed to persist funds", ex);
        }
    }

    @Override
    public void importBudgets(List<SclxDocument.Budget> budgets)
    {
        for (SclxDocument.Budget budget : budgets)
        {
            BudgetRecord row = new BudgetRecord(
                budget.budgetId(),
                budget.name(),
                budget.fiscalYear(),
                budget.fundId(),
                budget.active(),
                budget.description(),
                budget.lines() == null ? List.of() : budget.lines().stream()
                    .map(line -> new BudgetRecord.BudgetLineRecord(
                        line.eventName(),
                        line.budgetedAmount(),
                        line.revenueCategory(),
                        line.expenseCategory(),
                        resolveAccountNumber(line.accountId(), line.accountId()),
                        line.notes(),
                        line.extensions()))
                    .toList(),
                budget.extensions(),
                toJson(budget)
            );
            try
            {
                budgetRecordRepository.upsert(row);
            }
            catch (SQLException ex)
            {
                throw new IllegalStateException("Failed to persist budget " + budget.budgetId(), ex);
            }
        }
    }

    @Override
    public void importPeople(List<SclxDocument.Person> people)
    {
        for (SclxDocument.Person source : people)
        {
            if (source == null)
            {
                continue;
            }
            Person person = resolvePerson(source);
            if (source.personId() != null)
            {
                if (person.getId() > 0)
                {
                    personDbIdBySclxPersonId.put(source.personId(), person.getId());
                }
                personDisplayNameBySclxPersonId.put(source.personId(), source.displayName());
            }
        }
    }

    @Override
    public void importEvents(List<SclxDocument.Event> events)
    {
        upsertDocumentJson("sclx.events", events);
    }

    @Override
    public void importDocuments(List<SclxDocument.Document> documents)
    {
        upsertDocumentJson("sclx.documents", documents);
    }

    @Override
    public void importTransactions(List<SclxDocument.Transaction> transactions)
    {
        for (SclxDocument.Transaction source : transactions)
        {
            AccountingTransaction txn = mapTransaction(source);
            AccountingTransaction saved = journalGateway.saveTransactionWithEntries(txn);
            importedTransactionIdsBySclxId.put(source.transactionId(), saved.getId());
            if (source.workbookLink() != null && source.workbookLink().ledgerRowIndex() != null)
            {
                importedTransactionIdsByLedgerRow.put(source.workbookLink().ledgerRowIndex(), saved.getId());
            }
        }
    }

    @Override
    public void importOutstandingItems(List<SclxDocument.OutstandingItem> outstandingItems)
    {
        upsertDocumentJson("sclx.outstandingItems", outstandingItems);
    }

    @Override
    public void importOtherAssetItems(List<SclxDocument.OtherAssetItem> otherAssetItems)
    {
        upsertDocumentJson("sclx.otherAssetItems", otherAssetItems);
    }

    @Override
    public void importSupplementalItems(List<SclxDocument.SupplementalItem> supplementalItems)
    {
        for (SclxDocument.SupplementalItem item : supplementalItems)
        {
            Integer txnId = item.ledgerRowIndex() == null ? null : importedTransactionIdsByLedgerRow.get(item.ledgerRowIndex());
            if (txnId == null)
            {
                continue;
            }

            TxnSupplementalLineBase bean = createSupplementalBean(item.kind());
            bean.setTxnId(txnId.longValue());
            bean.setCounterpartyPersonId(resolveCounterpartyId(item.personId(), item.counterpartyName()));
            bean.setDescription(item.reason());
            bean.setReference(firstNonBlank(item.eventBudgetLabel(), item.sourceLabel(), item.subtypeCode()));
            bean.setAmount(item.amountAsOf());
            bean.setNotes(buildSupplementalNotes(item));

            TxnSupplementalLineRecord record = TxnSupplementalLineMapper.toRecord(bean);

            try (Connection c = Database.get().getConnection())
            {
                List<TxnSupplementalLineRecord> all = supplementalRepository.listByTxnId(txnId.longValue());
                all.add(record);
                c.setAutoCommit(false);
                supplementalRepository.replaceForTxn(c, txnId.longValue(), all);
                c.commit();
            }
            catch (SQLException ex)
            {
                throw new IllegalStateException("Failed to persist supplemental item " + item.supplementalItemId(), ex);
            }
        }
    }

    @Override
    public void importAssets(List<SclxDocument.Asset> assets)
    {
        upsertDocumentJson("sclx.assets", assets);
    }

    @Override
    public void importSupplies(List<SclxDocument.Supply> supplies)
    {
        upsertDocumentJson("sclx.supplies", supplies);
    }

    @Override
    public void importBankingItems(List<SclxDocument.BankingItem> bankingItems)
    {
        for (SclxDocument.BankingItem item : bankingItems)
        {
            BankingItemRecord row = new BankingItemRecord(
                item.bankingItemId(),
                item.kind(),
                item.bankAccountId(),
                item.transactionId(),
                List.of(),
                item.ofx() != null && item.ofx().datePosted() != null ? item.ofx().datePosted() : item.depositDate(),
                item.amount(),
                item.checkNumber(),
                item.payee(),
                item.depositDate(),
                item.payer(),
                null,
                item.ofx() == null ? null : item.ofx().memo(),
                item.source(),
                item.status(),
                item.importId(),
                item.ofx() == null ? null : new BankingItemRecord.OfxTransactionRecord(
                    item.ofx().fitId(),
                    item.ofx().transactionType(),
                    item.ofx().datePosted(),
                    item.ofx().dateUser(),
                    item.ofx().dateAvailable(),
                    item.ofx().checkNumber(),
                    item.ofx().referenceNumber(),
                    item.ofx().name(),
                    item.ofx().memo(),
                    item.ofx().payeeId(),
                    item.ofx().sic(),
                    item.ofx().serverTransactionId(),
                    item.ofx().correctFitId(),
                    item.ofx().correctAction(),
                    item.ofx().extensions()
                ),
                item.extensions(),
                toJson(item)
            );
            try
            {
                bankingItemRecordRepository.upsert(row);
            }
            catch (SQLException ex)
            {
                throw new IllegalStateException("Failed to persist banking item " + item.bankingItemId(), ex);
            }
        }
    }

    @Override
    public void importBankStatementImports(List<SclxDocument.BankStatementImport> bankStatementImports)
    {
        for (SclxDocument.BankStatementImport item : bankStatementImports)
        {
            BankStatementRecord row = new BankStatementRecord(
                item.importId(),
                parseSourceFormat(item.sourceFormat()),
                item.sourceVersion(),
                parseStatementKind(item.statementKind()),
                item.bankAccount() == null ? null : new BankStatementRecord.BankAccountRef(
                    item.bankAccount().bankId(),
                    item.bankAccount().accountId(),
                    item.bankAccount().accountType()
                ),
                item.currency(),
                item.statementStart(),
                item.statementEnd(),
                item.ledgerBalance() == null ? null : new BankStatementRecord.BalanceSnapshot(item.ledgerBalance().amount(), item.ledgerBalance().asOf()),
                item.availableBalance() == null ? null : new BankStatementRecord.BalanceSnapshot(item.availableBalance().amount(), item.availableBalance().asOf()),
                item.documentId(),
                item.extensions(),
                toJson(item)
            );
            try
            {
                bankStatementRecordRepository.upsert(row);
            }
            catch (SQLException ex)
            {
                throw new IllegalStateException("Failed to persist bank statement import " + item.importId(), ex);
            }
        }
    }

    @Override
    public void completeImport(SclxImportResult result)
    {
        upsertDocumentJson("sclx.importSummary", result);
    }

    
private AccountingTransaction mapTransaction(SclxDocument.Transaction source)
{
    LinkedHashSet<AccountingEntry> entries = new LinkedHashSet<>();
    for (SclxDocument.TransactionLine line : source.lines())
    {
        BigDecimal amount = debitAmount(line.debit(), line.credit());
        AccountSide side = line.debit() != null && line.debit().compareTo(BigDecimal.ZERO) > 0 ? AccountSide.DEBIT : AccountSide.CREDIT;
        AccountingEntry entry = new AccountingEntry(
            amount,
            resolveAccountNumber(line.accountId(), line.accountId()),
            side,
            line.accountId()
        );
        entry.setFundNumber(line.fundId());
        entries.add(entry);
    }

    BigDecimal debitTotal = entries.stream()
        .filter(e -> e.getAccountSide() == AccountSide.DEBIT)
        .map(AccountingEntry::getAmount)
        .reduce(BigDecimal.ZERO, BigDecimal::add);
    BigDecimal creditTotal = entries.stream()
        .filter(e -> e.getAccountSide() == AccountSide.CREDIT)
        .map(AccountingEntry::getAmount)
        .reduce(BigDecimal.ZERO, BigDecimal::add);
    BigDecimal delta = debitTotal.subtract(creditTotal);

    if (delta.compareTo(BigDecimal.ZERO) != 0)
    {
        if (!currentOptions.allowSingleSidedTransactions())
        {
            throw new IllegalStateException("Unbalanced SCLX transaction " + source.transactionId());
        }
        if (!currentOptions.hasCashAccountReference())
        {
            throw new IllegalStateException("cashAccountReference is required to import single-sided or unbalanced SCLX transactions.");
        }

        if (delta.compareTo(BigDecimal.ZERO) > 0)
        {
            entries.add(new AccountingEntry(delta.abs(), currentOptions.cashAccountReference(), AccountSide.CREDIT, "Cash"));
        }
        else
        {
            entries.add(new AccountingEntry(delta.abs(), currentOptions.cashAccountReference(), AccountSide.DEBIT, "Cash"));
        }
    }

    AccountingTransaction txn = new AccountingTransaction();
    txn.setEntries(entries);
    txn.setDate(source.transactionDate() == null ? (source.postingDate() == null ? "" : source.postingDate().toString()) : source.transactionDate().toString());
    txn.setMemo(source.description());
    txn.setToFrom(resolveTransactionPersonDisplayName(source));
    txn.setCheckNumber(firstNonBlank(source.checkNumber(), source.checkNumberId()));
    txn.setClearBank(source.bankTiming());
    txn.setBank(source.bankTiming());
    txn.setBudgetTracking(source.budgetId());
    txn.setAssociatedFundName(firstLineFund(source));
    txn.setBookingDateTimestamp(source.postingDate() == null ? System.currentTimeMillis() : source.postingDate().atStartOfDay().toInstant(ZoneOffset.UTC).toEpochMilli());

    Map<String, String> info = new LinkedHashMap<>();
    info.put("sclx.transactionId", nullToEmpty(source.transactionId()));
    info.put("sclx.source", nullToEmpty(source.source()));
    info.put("sclx.status", nullToEmpty(source.status()));
    info.put("sclx.checkNumberId", nullToEmpty(source.checkNumberId()));
    info.put("sclx.personId", nullToEmpty(source.personId()));
    info.put("sclx.personDisplayName", nullToEmpty(source.personDisplayName()));
    if (source.workbookLink() != null)
    {
        info.put("sclx.sheetKey", nullToEmpty(source.workbookLink().sheetKey()));
        if (source.workbookLink().ledgerRowIndex() != null)
        {
            info.put("sclx.ledgerRowIndex", String.valueOf(source.workbookLink().ledgerRowIndex()));
        }
    }
    txn.setInfo(info);
    txn.setSupplementalLines(new ArrayList<>());
    return txn;
}

private Person resolvePerson(SclxDocument.Person source)
    {
        try
        {
            for (Person existing : personRepository.list())
            {
                if (equalsIgnoreCase(existing.getName(), source.displayName()) ||
                    (source.email() != null && equalsIgnoreCase(existing.getEmail(), source.email())))
                {
                    if (source.email() != null && (existing.getEmail() == null || existing.getEmail().isBlank()))
                    {
                        existing.setEmail(source.email());
                    }
                    if (source.phone() != null && (existing.getPhone() == null || existing.getPhone().isBlank()))
                    {
                        existing.setPhone(source.phone());
                    }
                    return personRepository.save(existing);
                }
            }

            Person person = new Person();
            person.setName(source.displayName());
            person.setEmail(source.email());
            person.setPhone(source.phone());
            return personRepository.save(person);
        }
        catch (SQLException ex)
        {
            throw new IllegalStateException("Failed to resolve person " + source.displayName(), ex);
        }
    }

    private Long resolveCounterpartyId(String sclxPersonId, String counterpartyName)
    {
        if (sclxPersonId != null && personDbIdBySclxPersonId.containsKey(sclxPersonId))
        {
            return personDbIdBySclxPersonId.get(sclxPersonId);
        }
        if (counterpartyName == null || counterpartyName.isBlank())
        {
            return null;
        }
        try
        {
            for (Person existing : personRepository.list())
            {
                if (equalsIgnoreCase(existing.getName(), counterpartyName))
                {
                    return existing.getId();
                }
            }
            Person person = new Person();
            person.setName(counterpartyName);
            return personRepository.save(person).getId();
        }
        catch (SQLException ex)
        {
            throw new IllegalStateException("Failed to resolve counterparty " + counterpartyName, ex);
        }
    }

    private TxnSupplementalLineBase createSupplementalBean(String kind)
    {
        return switch (kind)
        {
            case "RECEIVABLE" -> new ReceivablesLine();
            case "PREPAID_EXPENSE" -> new PrepaidExpenseLine();
            case "OTHER_ASSET" -> new OtherAssetsLine();
            case "DEFERRED_REVENUE" -> new DeferredRevenueLine();
            case "PAYABLE" -> new PayablesLine();
            case "OTHER_LIABILITY" -> new OtherLiabilitiesLine();
            default -> throw new IllegalArgumentException("Unsupported supplemental kind: " + kind);
        };
    }

    private String buildSupplementalNotes(SclxDocument.SupplementalItem item)
    {
        List<String> parts = new ArrayList<>();
        if (item.subtypeCode() != null) parts.add("subtype=" + item.subtypeCode());
        if (item.eventBudgetLabel() != null) parts.add("budget=" + item.eventBudgetLabel());
        if (item.sourceLabel() != null) parts.add("source=" + item.sourceLabel());
        return String.join("; ", parts);
    }

    private SupplementalLineKind parseSupplementalKind(String value)
    {
        if (value == null || value.isBlank()) return null;
        try { return SupplementalLineKind.valueOf(value); }
        catch (IllegalArgumentException ex) { return null; }
    }

    private String resolveAccountNumber(String preferred, String fallback)
    {
        return currentOptions.resolveAccountReference(firstNonBlank(preferred, fallback));
    }

    private BigDecimal debitAmount(BigDecimal debit, BigDecimal credit)
    {
        if (debit != null && debit.compareTo(BigDecimal.ZERO) > 0) return debit;
        if (credit != null && credit.compareTo(BigDecimal.ZERO) > 0) return credit;
        return BigDecimal.ZERO;
    }

    private String firstLineFund(SclxDocument.Transaction source)
    {
        if (source.lines() == null) return "";
        for (SclxDocument.TransactionLine line : source.lines())
        {
            if (line.fundId() != null && !line.fundId().isBlank())
            {
                return line.fundId();
            }
        }
        return "";
    }

    private String firstNonBlank(String... values)
    {
        if (values == null) return null;
        for (String value : values)
        {
            if (value != null && !value.isBlank()) return value;
        }
        return null;
    }

    private boolean equalsIgnoreCase(String a, String b)
    {
        return a != null && b != null && a.equalsIgnoreCase(b);
    }

    private String nullToEmpty(String value)
    {
        return value == null ? "" : value;
    }

    private void upsertDocumentJson(String name, Object value)
    {
        try
        {
            documentRepository.upsert(name, toJson(value));
        }
        catch (SQLException ex)
        {
            throw new IllegalStateException("Failed to store document " + name, ex);
        }
    }

    private String toJson(Object value)
    {
        try
        {
            return MAPPER.writeValueAsString(value);
        }
        catch (JsonProcessingException ex)
        {
            throw new IllegalStateException("Failed to serialize value to JSON", ex);
        }
    }

    private BankStatementRecord.SourceFormat parseSourceFormat(String value)
    {
        if (value == null || value.isBlank()) return null;
        try { return BankStatementRecord.SourceFormat.valueOf(value); }
        catch (IllegalArgumentException ex) { return BankStatementRecord.SourceFormat.OTHER; }
    }

    private BankStatementRecord.StatementKind parseStatementKind(String value)
    {
        if (value == null || value.isBlank()) return null;
        try { return BankStatementRecord.StatementKind.valueOf(value); }
        catch (IllegalArgumentException ex) { return BankStatementRecord.StatementKind.OTHER; }
    }
}
