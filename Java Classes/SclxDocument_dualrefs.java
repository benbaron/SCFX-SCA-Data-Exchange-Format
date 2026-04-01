package nonprofitbookkeeping.importer.sclx;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.OffsetDateTime;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

/**
 * Root DTO and nested DTO types for SCLX import.
 *
 * <p>This model is intentionally tolerant:
 * <ul>
 *   <li>unknown properties are ignored</li>
 *   <li>ledger-native single-line transactions are allowed</li>
 *   <li>nullable outstanding ledger links are allowed</li>
 *   <li>supplementalItems are supported when present</li>
 * </ul>
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record SclxDocument(
    String format,
    String version,
    OffsetDateTime exportedAt,
    Organization organization,
    ReportingPeriod reportingPeriod,
    List<Account> chartOfAccounts,
    List<Fund> funds,
    List<Budget> budgets,
    List<Person> people,
    List<Event> events,
    List<Document> documents,
    List<Transaction> transactions,
    List<BankingItem> bankingItems,
    List<OutstandingItem> outstandingItems,
    List<OtherAssetItem> otherAssetItems,
    List<SupplementalItem> supplementalItems,
    List<Asset> assets,
    List<Supply> supplies,
    List<BankStatementImport> bankStatementImports,
    Map<String, Object> extensions)
{
    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Organization(
        String organizationId,
        String name,
        String parentOrganization,
        String baseCurrency,
        LocalDate fiscalYearStart,
        LocalDate fiscalYearEnd,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record ReportingPeriod(
        LocalDate startDate,
        LocalDate endDate,
        String label,
        Integer fiscalYear,
        String periodType,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Account(
        String Number,
        String Name,
        String Type,
        String Parent,
        String IncreaseSide,
        BigDecimal OpeningBalance,
        List<String> SupplementalKinds,
        String accountId,
        String code,
        String subtype,
        Boolean active,
        List<String> reportingTags,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Fund(
        String fundId,
        String name,
        Boolean restricted,
        String description,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Budget(
        String budgetId,
        String name,
        Integer fiscalYear,
        String fundId,
        Boolean active,
        String description,
        List<BudgetLine> lines,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record BudgetLine(
        String eventName,
        BigDecimal budgetedAmount,
        String revenueCategory,
        String expenseCategory,
        String accountId,
        String notes,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Person(
        String personId,
        String displayName,
        String kind,
        String email,
        String phone,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Event(
        String eventId,
        String name,
        LocalDate startDate,
        LocalDate endDate,
        String hostingOrganizationId,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Document(
        String documentId,
        String documentType,
        String referenceNumber,
        LocalDate documentDate,
        String fileName,
        String notes,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Transaction(
        String transactionId,
        LocalDate transactionDate,
        LocalDate postingDate,
        String description,
        String checkNumber,
        String checkNumberId,
        String personId,
        String personDisplayName,
        String status,
        String source,
        String bankTiming,
        String budgetTiming,
        String budgetId,
        WorkbookLink workbookLink,
        Approval approval,
        List<String> documentIds,
        String eventId,
        List<TransactionLine> lines,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record TransactionLine(
        String lineId,
        String accountId,
        BigDecimal debit,
        BigDecimal credit,
        String fundId,
        String budgetId,
        String personId,
        String eventId,
        String documentId,
        String memo,
        List<String> tags,
        String restrictionTag,
        String reportSection,
        String usedFor,
        String itemNumber,
        Integer quantity,
        WorkbookLink workbookLink,
        List<SupplementalRef> supplementalRefs,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Approval(
        Boolean policyRequired,
        String committeeApprovalRef,
        List<String> approvedBy,
        LocalDate approvalDate,
        String notes,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record LedgerLink(
        String transactionId,
        String lineId)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record WorkbookLink(
        String sheetKey,
        Integer ledgerRowIndex)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record SupplementalRef(
        String recordType,
        String recordId)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record OutstandingItem(
        String outstandingItemId,
        String kind,
        LedgerLink ledgerLink,
        WorkbookLink workbookLink,
        LocalDate dateSentOrReceived,
        LocalDate incomingCheckOrTransferDate,
        String transferIdOrCheckNumber,
        LocalDate dateShowsOnStatement,
        String personOrBusinessName,
        String detailsNotes,
        String fromToCardMerchant,
        String accountForPaymentOrDeposit,
        BigDecimal amount,
        LocalDate dateReversed,
        String reversalReasonAndApproval,
        LedgerLink reversalLedgerLink,
        String status,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record OtherAssetItem(
        String otherAssetItemId,
        LedgerLink ledgerLink,
        WorkbookLink workbookLink,
        String paidTo,
        Integer year,
        String reason,
        String type,
        String typeCode,
        String eventBudgetLabel,
        BigDecimal amountAsOfPriorYearEnd,
        Integer paidReturnedOnLedgerRowIndex,
        LedgerLink settlementLedgerLink,
        String status,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record SupplementalItem(
        String supplementalItemId,
        String kind,
        String counterpartyName,
        String personId,
        Integer year,
        String reason,
        String subtypeCode,
        @JsonProperty("eventBudgetLabel") String eventBudgetLabel,
        String budgetId,
        String sourceLabel,
        BigDecimal amountAsOf,
        Integer ledgerRowIndex,
        WorkbookLink workbookLink,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Asset(
        String assetId,
        LocalDate dateAcquired,
        String description,
        Integer itemCount,
        BigDecimal approxValueTotal,
        BigDecimal valuePerItem,
        String itemType,
        String usedFor,
        BigDecimal lotPaidTotal,
        Integer lotItemCount,
        Guardian currentGuardian,
        GuardianshipDetailsAsset guardianshipDetails,
        RemovalDetailsAsset removalDetails,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Guardian(
        String legalName,
        String email,
        String phone)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record GuardianshipDetailsAsset(
        LocalDate dateAsOf,
        Boolean confirmed,
        String confirmationStatus,
        String notes)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record GuardianshipDetailsSupply(
        LocalDate dateAsOf,
        LocalDate lastConfirmed,
        Boolean returned,
        String notes)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record RemovalDetailsAsset(
        String approvedBy,
        LocalDate approvalDate,
        String reason,
        Integer numberRemoved,
        Boolean removed,
        String removalType)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record RemovalDetailsSupply(
        String approvedBy,
        String reason,
        Integer numberRemoved,
        Boolean removed,
        String removalType)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Supply(
        String supplyId,
        String itemNumber,
        LocalDate dateAcquired,
        String description,
        Integer count,
        BigDecimal approxValueTotal,
        BigDecimal valuePerItem,
        Guardian guardian,
        GuardianshipDetailsSupply guardianshipDetails,
        RemovalDetailsSupply removalDetails,
        String additionalNotes,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record BankStatementImport(
        String importId,
        String sourceFormat,
        String sourceVersion,
        String statementKind,
        BankAccount bankAccount,
        String currency,
        LocalDate statementStart,
        LocalDate statementEnd,
        StatementBalance ledgerBalance,
        StatementBalance availableBalance,
        String documentId,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record BankAccount(
        String bankId,
        String accountId,
        String accountType)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record StatementBalance(
        BigDecimal amount,
        OffsetDateTime asOf)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record BankingItem(
        String bankingItemId,
        String kind,
        String source,
        String status,
        String transactionId,
        String bankAccountId,
        LocalDate depositDate,
        String payer,
        String payee,
        String checkNumber,
        BigDecimal amount,
        String importId,
        OfxTransaction ofx,
        Map<String, Object> extensions)
    {
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record OfxTransaction(
        String fitId,
        String transactionType,
        LocalDate datePosted,
        LocalDate dateUser,
        LocalDate dateAvailable,
        String checkNumber,
        String referenceNumber,
        String name,
        String memo,
        String payeeId,
        String sic,
        String serverTransactionId,
        String correctFitId,
        String correctAction,
        Map<String, Object> extensions)
    {
    }
}
