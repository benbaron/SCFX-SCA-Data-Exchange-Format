Attribute VB_Name = "SCLX_Ledger_IO"
Option Explicit

'===============================================================================
' SCLX_Ledger_IO.bas
'
' Purpose
'   Export/import SCLX JSON to/from the "SCA Exchequer Report - 2026-03.xlsx"
'   workbook layout.
'
' Dependencies
'   1) Import JsonConverter.bas from VBA-JSON
'   2) No compile-time reference is required; dictionaries are late-bound.
'
' Notes
'   If you paste this directly into the VBA editor, remove the Attribute line.
'===============================================================================

Private Const SCLX_VERSION As String = "1.3"

Private Const SH_SUMMARY As String = "Summary"
Private Const SH_LEDGER As String = "Ledger"
Private Const SH_OUTSTANDING As String = "Outstanding"
Private Const SH_ASSETS As String = "Assets&Inventory"
Private Const SH_SUPPLIES As String = "Supplies"

Private Const ROW_LEDGER_FIRST As Long = 9
Private Const ROW_OUT_FIRST As Long = 14
Private Const ROW_ASSET_FIRST As Long = 11
Private Const ROW_SUPPLY_FIRST As Long = 10

Private Const CELL_SUM_ORG_NAME As String = "D9"
Private Const CELL_SUM_PARENT_ORG As String = "D7"
Private Const CELL_SUM_CURRENCY As String = "H8"
Private Const CELL_SUM_LEDGER_YEAR As String = "H6"
Private Const CELL_SUM_REPORT_QTR As String = "H7"
Private Const CELL_SUM_REPORT_LABEL As String = "I3"
Private Const CELL_SUM_REPORT_START As String = "T5"
Private Const CELL_SUM_REPORT_END As String = "T6"
Private Const CELL_SUM_FY_END As String = "T8"

Private mBudgetNameById As Object
Private mFundNameById As Object

' Ledger visible transaction columns
Private Const COL_LEDGER_ROWNUM As String = "A"
Private Const COL_LEDGER_TXN_DATE As String = "D"
Private Const COL_LEDGER_DATE_SHOWS As String = "E"
Private Const COL_LEDGER_REF As String = "F"
Private Const COL_LEDGER_INCOMING As String = "G"
Private Const COL_LEDGER_NAME As String = "H"
Private Const COL_LEDGER_DETAILS As String = "I"
Private Const COL_LEDGER_BANK_ACCOUNT As String = "J"
Private Const COL_LEDGER_AFFECTS_BANK As String = "K"
Private Const COL_LEDGER_BUDGET_CATEGORY As String = "L"
Private Const COL_LEDGER_AFFECTS_BUDGET As String = "M"
Private Const COL_LEDGER_FUND As String = "N"
Private Const COL_LEDGER_MERCHANT As String = "O"

' Ledger split entry blocks (1..4)
' Each group is:
'   0 Split Row Number
'   1 Amount
'   2 Income Category
'   3 Expense Category
'   4 Used For
'   5 Item Num
'   6 Qty
Private ledgerSplitCols As Variant

' Outstanding columns
Private Const COL_OUT_OSROW As String = "B"
Private Const COL_OUT_DATE_SENT As String = "C"
Private Const COL_OUT_INCOMING_DATE As String = "D"
Private Const COL_OUT_TRANSFER_OR_CHECK As String = "E"
Private Const COL_OUT_DATE_SHOWS As String = "F"
Private Const COL_OUT_NAME As String = "G"
Private Const COL_OUT_DETAILS As String = "H"
Private Const COL_OUT_MERCHANT As String = "I"
Private Const COL_OUT_ACCOUNT As String = "J"
Private Const COL_OUT_AMOUNT As String = "K"
Private Const COL_OUT_DATE_REVERSED As String = "L"
Private Const COL_OUT_REASON_APPROVAL As String = "M"

' Assets&Inventory columns
Private Const COL_ASSET_ITEMNUM As String = "A"
Private Const COL_ASSET_DATE_ACQ As String = "B"
Private Const COL_ASSET_DESC As String = "C"
Private Const COL_ASSET_ITEM_COUNT As String = "D"
Private Const COL_ASSET_TOTAL_VALUE As String = "E"
Private Const COL_ASSET_TOTAL_LOT_COUNT As String = "F"
Private Const COL_ASSET_TOTAL_PAID As String = "G"
Private Const COL_ASSET_PER_ITEM As String = "H"
Private Const COL_ASSET_ITEM_TYPE As String = "I"
Private Const COL_ASSET_USED_FOR As String = "J"
Private Const COL_ASSET_GUARDIAN_NAME As String = "K"
Private Const COL_ASSET_GUARDIAN_EMAIL As String = "L"
Private Const COL_ASSET_GUARDIAN_PHONE As String = "M"
Private Const COL_ASSET_DATE_AS_OF As String = "N"
Private Const COL_ASSET_CONFIRMED As String = "O"
Private Const COL_ASSET_NOTES As String = "P"
Private Const COL_ASSET_APPROVED_BY As String = "Q"
Private Const COL_ASSET_DATE_REMOVED As String = "R"
Private Const COL_ASSET_REASON As String = "S"
Private Const COL_ASSET_NUM_REMOVED As String = "T"

' Supplies columns
Private Const COL_SUP_ITEMNUM As String = "A"
Private Const COL_SUP_DATE_ACQ As String = "B"
Private Const COL_SUP_DESC As String = "C"
Private Const COL_SUP_COUNT As String = "D"
Private Const COL_SUP_TOTAL_VALUE As String = "E"
Private Const COL_SUP_PER_ITEM As String = "F"
Private Const COL_SUP_GUARDIAN_NAME As String = "G"
Private Const COL_SUP_GUARDIAN_EMAIL As String = "H"
Private Const COL_SUP_GUARDIAN_PHONE As String = "I"
Private Const COL_SUP_DATE_AS_OF As String = "J"
Private Const COL_SUP_LAST_CONFIRMED As String = "K"
Private Const COL_SUP_RETURNED As String = "L"
Private Const COL_SUP_NOTES As String = "M"
Private Const COL_SUP_APPROVED_BY As String = "N"
Private Const COL_SUP_REASON As String = "O"
Private Const COL_SUP_NUMBER_REMOVED As String = "P"
Private Const COL_SUP_ADDITIONAL_NOTES As String = "Q"


Public Sub ExportSCLX()
    On Error GoTo EH

    InitSplitCols

    Dim path As Variant

    path = Application.GetSaveAsFilename(InitialFilename:="ledger.sclx.json", fileFilter:="JSON Files (*.json), *.json")

    If path <> False Then
        MsgBox "Save as " & path
    End If
    If path = False Then Exit Sub

    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")

    root("format") = "SCLX"
    root("version") = SCLX_VERSION
    root("exportedAt") = FormatDateTimeOffset(Now)
    root.Add "organization", ExportOrganization()
    root.Add "reportingPeriod", ExportReportingPeriod()
    root.Add "chartOfAccounts", ExportChartOfAccounts()
    root.Add "funds", ExportFunds()
    root.Add "budgets", ExportBudgets()
    root.Add "people", NewJsonArray()
    root.Add "events", NewJsonArray()
    root.Add "documents", NewJsonArray()
    root.Add "transactions", ExportTransactions()
    root.Add "bankingItems", NewJsonArray()
    root.Add "outstandingItems", ExportOutstandingItems()
    root.Add "otherAssetItems", NewJsonArray()
    root.Add "assets", ExportAssets()
    root.Add "supplies", ExportSupplies()
    root.Add "bankStatementImports", NewJsonArray()
    root.Add "extensions", CreateObject("Scripting.Dictionary")

    WriteTextFile CStr(path), JsonConverter.ConvertToJson(root, Whitespace:=2)
    MsgBox "SCLX export written to:" & vbCrLf & CStr(path), vbInformation
    Exit Sub

EH:
    MsgBox "ExportSCLX failed: " & Err.Description, vbCritical
End Sub


Public Sub ImportSCLX()
    On Error GoTo EH

    InitSplitCols
    ClearImportLookupMaps

    Dim path As Variant
    path = Application.GetOpenFilename(FileFilter:="JSON Files (*.json), *.json")
    If VarType(path) = vbBoolean Then Exit Sub

    Dim jsonText As String
    jsonText = ReadTextFile(CStr(path))

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)

    If Not root.Exists("format") Or CStr(root("format")) <> "SCLX" Then
        Err.Raise vbObjectError + 1000, , "File is not SCLX."
    End If

    BuildImportLookupMaps root

    If MsgBox("This will append SCLX data into the workbook tabs." & vbCrLf & _
              "Continue?", vbQuestion + vbOKCancel) <> vbOK Then Exit Sub

    If root.Exists("transactions") Then ImportTransactions root("transactions")
    If root.Exists("outstandingItems") Then ImportOutstandingItems root("outstandingItems")
    If root.Exists("assets") Then ImportAssets root("assets")
    If root.Exists("supplies") Then ImportSupplies root("supplies")

    MsgBox "SCLX import completed.", vbInformation
    Exit Sub

EH:
    MsgBox "ImportSCLX failed: " & Err.Description, vbCritical
End Sub


Private Function ExportOrganization() As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SUMMARY)

    Dim d As Object
    Dim ext As Object
    Dim nm As String
    Dim fyStart As String
    Dim fyEnd As String

    nm = SafeText(ws.Range(CELL_SUM_ORG_NAME).Value)
    If Len(nm) = 0 Then nm = SafeText(ws.Range("B2").Value)

    fyStart = Format$(DateSerial(SummaryYear(ws), 1, 1), "yyyy-mm-dd")
    If IsDate(ws.Range(CELL_SUM_FY_END).Value) Then
        fyEnd = Format$(CDate(ws.Range(CELL_SUM_FY_END).Value), "yyyy-mm-dd")
    Else
        fyEnd = Format$(DateSerial(SummaryYear(ws), 12, 31), "yyyy-mm-dd")
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d("organizationId") = NormalizeId("org-", nm)
    d("name") = nm
    d("parentOrganization") = SafeOrNull(ws.Range(CELL_SUM_PARENT_ORG).Value)
    d("baseCurrency") = SafeText(ws.Range(CELL_SUM_CURRENCY).Value)
    d("fiscalYearStart") = fyStart
    d("fiscalYearEnd") = fyEnd

    Set ext = CreateObject("Scripting.Dictionary")
    ext("branchType") = SafeOrNull(ws.Range("D6").Value)
    ext("location") = SafeOrNull(ws.Range("D8").Value)
    d.Add "extensions", ext

    Set ExportOrganization = d
End Function


Private Function ExportReportingPeriod() As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SUMMARY)

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("startDate") = DateCellOrFallback(ws.Range(CELL_SUM_REPORT_START).Value, Format$(DateSerial(SummaryYear(ws), 1, 1), "yyyy-mm-dd"))
    d("endDate") = DateCellOrFallback(ws.Range(CELL_SUM_REPORT_END).Value, Format$(DateSerial(SummaryYear(ws), 12, 31), "yyyy-mm-dd"))
    d("label") = SafeText(ws.Range(CELL_SUM_REPORT_LABEL).Value)
    d("fiscalYear") = SummaryYear(ws)
    d("periodType") = "QUARTER"
    Set ExportReportingPeriod = d
End Function

Private Function ExportChartOfAccounts() As Collection
    Dim accounts As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim txs As Collection
    Dim t As Variant
    Dim lines As Object
    Dim line As Variant

    Set txs = ExportTransactions()

    For Each t In txs
        If ExistsInDict(t, "lines") Then
            Set lines = t("lines")
            For Each line In lines
                If ExistsInDict(line, "accountId") Then
                    AddSimpleAccount accounts, seen, CStr(line("accountId"))
                End If
            Next line
        End If
    Next t

    Dim wsL As Worksheet
    Dim wsO As Worksheet
    Dim r As Long

    Set wsL = ThisWorkbook.Worksheets(SH_LEDGER)
    For r = ROW_LEDGER_FIRST To FindLastInterestingLedgerRow(wsL)
        If IsLedgerRowUsed(wsL, r) Then
            AddSimpleAccount accounts, seen, SafeText(wsL.Cells(r, COL_LEDGER_BANK_ACCOUNT).Value)
            AddSimpleAccount accounts, seen, SafeText(wsL.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value)
        End If
    Next r

    Set wsO = ThisWorkbook.Worksheets(SH_OUTSTANDING)
    For r = ROW_OUT_FIRST To FindLastUsedByAnyValue(wsO, Array(COL_OUT_ACCOUNT, COL_OUT_NAME, COL_OUT_AMOUNT))
        If RowHasAnyValue(wsO, r, Array(COL_OUT_ACCOUNT, COL_OUT_NAME, COL_OUT_AMOUNT)) Then
            AddSimpleAccount accounts, seen, SafeText(wsO.Cells(r, COL_OUT_ACCOUNT).Value)
        End If
    Next r

    Set ExportChartOfAccounts = accounts
End Function

Private Function ExportFunds() As Collection
    Dim funds As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LEDGER)

    Dim r As Long
    Dim fundName As String
    Dim d As Object

    For r = ROW_LEDGER_FIRST To FindLastInterestingLedgerRow(ws)
        If IsLedgerRowUsed(ws, r) Then
            fundName = SafeText(ws.Cells(r, COL_LEDGER_FUND).Value)
            If Len(fundName) > 0 Then
                If Not seen.Exists(UCase$(fundName)) Then
                    seen.Add UCase$(fundName), True
                    Set d = CreateObject("Scripting.Dictionary")
                    d("fundId") = NormalizeId("fund-", fundName)
                    d("name") = fundName
                    d("restricted") = False
                    funds.Add d
                End If
            End If
        End If
    Next r

    Set ExportFunds = funds
End Function


Private Function ExportBudgets() As Collection
    Dim coll As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LEDGER)

    Dim r As Long
    Dim budgetName As String
    Dim fundName As String
    Dim budgetId As String
    Dim d As Object

    For r = ROW_LEDGER_FIRST To FindLastInterestingLedgerRow(ws)
        If IsLedgerRowUsed(ws, r) Then
            budgetName = SafeText(ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value)
            fundName = SafeText(ws.Cells(r, COL_LEDGER_FUND).Value)

            If Len(budgetName) > 0 And Len(fundName) > 0 Then
                budgetId = CStr(BudgetIdFromFields(budgetName, fundName))

                If Not seen.Exists(UCase$(budgetId)) Then
                    seen.Add UCase$(budgetId), True

                    Set d = CreateObject("Scripting.Dictionary")
                    d("budgetId") = budgetId
                    d("name") = budgetName
                    d("fiscalYear") = SummaryYear(ThisWorkbook.Worksheets(SH_SUMMARY))
                    d("fundId") = NormalizeId("fund-", fundName)
                    d("active") = True
                    d("description") = "Synthesized from Ledger budget category/fund values."
                    d.Add "lines", NewJsonArray()
                    d.Add "extensions", CreateObject("Scripting.Dictionary")
                    coll.Add d
                End If
            End If
        End If
    Next r

    Set ExportBudgets = coll
End Function

Private Function ExportTransactions() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LEDGER)

    Dim txs As New Collection
    Dim lastRow As Long
    Dim r As Long

    lastRow = FindLastInterestingLedgerRow(ws)

    For r = ROW_LEDGER_FIRST To lastRow
        If IsLedgerRowUsed(ws, r) Then
            txs.Add ExportLedgerRowAsTransaction(ws, r)
        End If
    Next r

    Set ExportTransactions = txs
End Function


Private Function ExportLedgerRowAsTransaction(ws As Worksheet, ByVal r As Long) As Object
    Dim tx As Object
    Dim ext As Object
    Dim wbk As Object
    Dim budgetId As Variant

    Set tx = CreateObject("Scripting.Dictionary")
    tx("transactionId") = "ledger-row-" & CStr(r)
    tx("transactionDate") = ISODateOrNull(ws.Cells(r, COL_LEDGER_TXN_DATE).Value)
    tx("postingDate") = ISODateOrNull(ws.Cells(r, COL_LEDGER_TXN_DATE).Value)
    tx("description") = SafeText(ws.Cells(r, COL_LEDGER_DETAILS).Value)
    tx("reference") = SafeText(ws.Cells(r, COL_LEDGER_REF).Value)
    tx("status") = "POSTED"
    tx("source") = "MANUAL"
    tx("bankTiming") = TimingValueFromWorkbook(ws.Cells(r, COL_LEDGER_AFFECTS_BANK).Value)
    tx("budgetTiming") = TimingValueFromWorkbook(ws.Cells(r, COL_LEDGER_AFFECTS_BUDGET).Value)

    budgetId = BudgetIdFromFields(SafeText(ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value), _
                                  SafeText(ws.Cells(r, COL_LEDGER_FUND).Value))
    If Not IsNull(budgetId) Then tx("budgetId") = budgetId

    tx.Add "workbookLink", WorkbookLinkObject(SH_LEDGER, r)

    Set ext = CreateObject("Scripting.Dictionary")
    Set wbk = CreateObject("Scripting.Dictionary")

    wbk("sheet") = SH_LEDGER
    wbk("ledgerRow") = r
    wbk("visibleRowNumber") = SafeOrNull(ws.Cells(r, COL_LEDGER_ROWNUM).Value)
    wbk("dateShowsOnStatement") = ISODateOrNull(ws.Cells(r, COL_LEDGER_DATE_SHOWS).Value)
    wbk("incomingCheckOrTransferDate") = ISODateOrNull(ws.Cells(r, COL_LEDGER_INCOMING).Value)
    wbk("personOrBusinessName") = SafeText(ws.Cells(r, COL_LEDGER_NAME).Value)
    wbk("detailsNotes") = SafeText(ws.Cells(r, COL_LEDGER_DETAILS).Value)
    wbk("bankAccount") = SafeText(ws.Cells(r, COL_LEDGER_BANK_ACCOUNT).Value)
    wbk("affectsBank") = SafeText(ws.Cells(r, COL_LEDGER_AFFECTS_BANK).Value)
    wbk("budgetCategory") = SafeText(ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value)
    wbk("affectsBudget") = SafeText(ws.Cells(r, COL_LEDGER_AFFECTS_BUDGET).Value)
    wbk("fund") = SafeText(ws.Cells(r, COL_LEDGER_FUND).Value)
    wbk("merchant") = SafeText(ws.Cells(r, COL_LEDGER_MERCHANT).Value)

    ext.Add "workbook", wbk
    tx.Add "extensions", ext
    tx.Add "lines", ExportLedgerSplitLines(ws, r)

    Set ExportLedgerRowAsTransaction = tx
End Function

Private Function ExportLedgerSplitLines(ws As Worksheet, ByVal r As Long) As Collection
    Dim lines As New Collection
    Dim i As Long
    Dim grp As Variant

    For i = LBound(ledgerSplitCols) To UBound(ledgerSplitCols)
        grp = ledgerSplitCols(i)
        If HasSplitData(ws, r, grp) Then
            lines.Add ExportOneSplit(ws, r, grp, i + 1)
        End If
    Next i

    Set ExportLedgerSplitLines = lines
End Function


Private Function ExportOneSplit(ws As Worksheet, ByVal r As Long, grp As Variant, ByVal splitIndex As Long) As Object
    Dim d As Object
    Dim ext As Object
    Dim wbk As Object
    Dim amt As Double
    Dim incomeCat As String
    Dim expenseCat As String
    Dim acct As String
    Dim fundName As String
    Dim budgetId As Variant

    amt = CDbl(ValZero(ws.Cells(r, grp(1)).Value))
    incomeCat = SafeText(ws.Cells(r, grp(2)).Value)
    expenseCat = SafeText(ws.Cells(r, grp(3)).Value)
    fundName = SafeText(ws.Cells(r, COL_LEDGER_FUND).Value)

    Set d = CreateObject("Scripting.Dictionary")
    d("lineId") = "ledger-row-" & r & "-ln-" & splitIndex

    If Len(expenseCat) > 0 Then
        acct = expenseCat
        d("debit") = FormatAmount(Abs(amt))
        d("credit") = FormatAmount(0)
    Else
        acct = incomeCat
        d("debit") = FormatAmount(0)
        d("credit") = FormatAmount(Abs(amt))
    End If
    If Len(acct) = 0 Then acct = "UNMAPPED"

    d("accountId") = acct

    If Len(fundName) > 0 Then
        d("fundId") = NormalizeId("fund-", fundName)
    Else
        d("fundId") = Null
    End If

    budgetId = BudgetIdFromFields(SafeText(ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value), fundName)
    If Not IsNull(budgetId) Then
        d("budgetId") = budgetId
    Else
        d("budgetId") = Null
    End If

    d("personId") = SafeOrNull(ws.Cells(r, COL_LEDGER_NAME).Value)
    d("eventId") = Null
    d("documentId") = Null
    d("memo") = SafeText(ws.Cells(r, COL_LEDGER_DETAILS).Value)
    d("usedFor") = SafeOrNull(ws.Cells(r, grp(4)).Value)
    d("itemNumber") = SafeOrNull(ws.Cells(r, grp(5)).Value)
    d("quantity") = NullOrNumber(ws.Cells(r, grp(6)).Value)
    d.Add "workbookLink", WorkbookLinkObject(SH_LEDGER, r)

    Set ext = CreateObject("Scripting.Dictionary")
    Set wbk = CreateObject("Scripting.Dictionary")
    wbk("splitIndex") = splitIndex
    wbk("splitRowNumber") = SafeOrNull(ws.Cells(r, grp(0)).Value)
    wbk("amount") = FormatAmount(Abs(amt))
    wbk("incomeCategory") = NullIfEmpty(incomeCat)
    wbk("expenseCategory") = NullIfEmpty(expenseCat)
    wbk("usedFor") = SafeOrNull(ws.Cells(r, grp(4)).Value)
    wbk("itemNumber") = SafeOrNull(ws.Cells(r, grp(5)).Value)
    wbk("quantity") = NullOrNumber(ws.Cells(r, grp(6)).Value)
    ext.Add "workbook", wbk
    d.Add "extensions", ext

    Set ExportOneSplit = d
End Function


Private Function ExportOutstandingItems() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_OUTSTANDING)

    Dim coll As New Collection
    Dim lastRow As Long
    Dim r As Long
    Dim d As Object
    Dim ext As Object
    Dim wbk As Object

    lastRow = FindLastInterestingOutstandingRow(ws)

    For r = ROW_OUT_FIRST To lastRow
        If IsOutstandingRowUsed(ws, r) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("outstandingItemId") = "outstanding-row-" & r
            d("kind") = GuessOutstandingKind(ws, r)
            d("ledgerLink") = Null
            d.Add "workbookLink", WorkbookLinkObject(SH_OUTSTANDING, r)
            d("amount") = FormatAmountAbs(ws.Cells(r, COL_OUT_AMOUNT).Value)

            Set ext = CreateObject("Scripting.Dictionary")
            Set wbk = CreateObject("Scripting.Dictionary")
            wbk("sheet") = SH_OUTSTANDING
            wbk("row") = r
            wbk("visibleRowNumber") = SafeOrNull(ws.Cells(r, COL_OUT_OSROW).Value)
            ext.Add "workbook", wbk
            d.Add "extensions", ext

            d("dateSentOrReceived") = ISODateOrNull(ws.Cells(r, COL_OUT_DATE_SENT).Value)
            d("incomingCheckOrTransferDate") = ISODateOrNull(ws.Cells(r, COL_OUT_INCOMING_DATE).Value)
            d("transferIdOrCheckNumber") = SafeOrNull(ws.Cells(r, COL_OUT_TRANSFER_OR_CHECK).Value)
            d("dateShowsOnStatement") = ISODateOrNull(ws.Cells(r, COL_OUT_DATE_SHOWS).Value)
            d("personOrBusinessName") = SafeOrNull(ws.Cells(r, COL_OUT_NAME).Value)
            d("detailsNotes") = SafeOrNull(ws.Cells(r, COL_OUT_DETAILS).Value)
            d("fromToCardMerchant") = SafeOrNull(ws.Cells(r, COL_OUT_MERCHANT).Value)
            d("accountForPaymentOrDeposit") = SafeOrNull(ws.Cells(r, COL_OUT_ACCOUNT).Value)
            d("dateReversed") = ISODateOrNull(ws.Cells(r, COL_OUT_DATE_REVERSED).Value)
            d("reversalReasonAndApproval") = SafeOrNull(ws.Cells(r, COL_OUT_REASON_APPROVAL).Value)
            d("status") = GuessOutstandingStatus(ws, r)
            coll.Add d
        End If
    Next r

    Set ExportOutstandingItems = coll
End Function


Private Function ExportAssets() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_ASSETS)

    Dim coll As New Collection
    Dim lastRow As Long
    Dim r As Long
    Dim d As Object

    lastRow = FindLastInterestingAssetRow(ws)

    For r = ROW_ASSET_FIRST To lastRow
        If IsAssetRowUsed(ws, r) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("assetId") = "asset-row-" & r
            d("dateAcquired") = ISODateOrNull(ws.Cells(r, COL_ASSET_DATE_ACQ).Value)
            d("description") = SafeOrNull(ws.Cells(r, COL_ASSET_DESC).Value)
            d("itemCount") = NullOrNumber(ws.Cells(r, COL_ASSET_ITEM_COUNT).Value)
            d("approxValueTotal") = AmountOrNull(ws.Cells(r, COL_ASSET_TOTAL_VALUE).Value)
            d("valuePerItem") = AmountOrNull(ws.Cells(r, COL_ASSET_PER_ITEM).Value)
            d("itemType") = SafeOrNull(ws.Cells(r, COL_ASSET_ITEM_TYPE).Value)
            d("usedFor") = SafeOrNull(ws.Cells(r, COL_ASSET_USED_FOR).Value)
            d("lotPaidTotal") = AmountOrNull(ws.Cells(r, COL_ASSET_TOTAL_PAID).Value)
            d("lotItemCount") = NullOrNumber(ws.Cells(r, COL_ASSET_TOTAL_LOT_COUNT).Value)
            d.Add "currentGuardian", GuardianObject(ws.Cells(r, COL_ASSET_GUARDIAN_NAME).Value, ws.Cells(r, COL_ASSET_GUARDIAN_EMAIL).Value, ws.Cells(r, COL_ASSET_GUARDIAN_PHONE).Value)
            d.Add "guardianshipDetails", GuardianshipObject(ws.Cells(r, COL_ASSET_DATE_AS_OF).Value, ws.Cells(r, COL_ASSET_CONFIRMED).Value, ws.Cells(r, COL_ASSET_NOTES).Value)
            d.Add "removalDetails", RemovalObject(ws.Cells(r, COL_ASSET_APPROVED_BY).Value, ws.Cells(r, COL_ASSET_DATE_REMOVED).Value, ws.Cells(r, COL_ASSET_REASON).Value, ws.Cells(r, COL_ASSET_NUM_REMOVED).Value)
            d.Add "extensions", WorkbookRowExtension(SH_ASSETS, r)
            coll.Add d
        End If
    Next r

    Set ExportAssets = coll
End Function


Private Function ExportSupplies() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SUPPLIES)

    Dim coll As New Collection
    Dim lastRow As Long
    Dim r As Long
    Dim d As Object
    Dim gd As Object

    lastRow = FindLastInterestingSupplyRow(ws)

    For r = ROW_SUPPLY_FIRST To lastRow
        If IsSupplyRowUsed(ws, r) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("supplyId") = "supply-row-" & r
            d("itemNumber") = SafeOrNull(ws.Cells(r, COL_SUP_ITEMNUM).Value)
            d("dateAcquired") = ISODateOrNull(ws.Cells(r, COL_SUP_DATE_ACQ).Value)
            d("description") = SafeOrNull(ws.Cells(r, COL_SUP_DESC).Value)
            d("count") = NullOrNumber(ws.Cells(r, COL_SUP_COUNT).Value)
            d("approxValueTotal") = AmountOrNull(ws.Cells(r, COL_SUP_TOTAL_VALUE).Value)
            d("valuePerItem") = AmountOrNull(ws.Cells(r, COL_SUP_PER_ITEM).Value)
            d.Add "guardian", GuardianObject(ws.Cells(r, COL_SUP_GUARDIAN_NAME).Value, ws.Cells(r, COL_SUP_GUARDIAN_EMAIL).Value, ws.Cells(r, COL_SUP_GUARDIAN_PHONE).Value)

            Set gd = CreateObject("Scripting.Dictionary")
            gd("dateAsOf") = ISODateOrNull(ws.Cells(r, COL_SUP_DATE_AS_OF).Value)
            gd("lastConfirmed") = ISODateOrNull(ws.Cells(r, COL_SUP_LAST_CONFIRMED).Value)
            gd("returned") = BoolOrNull(ws.Cells(r, COL_SUP_RETURNED).Value)
            gd("notes") = SafeOrNull(ws.Cells(r, COL_SUP_NOTES).Value)
            d.Add "guardianshipDetails", gd

            d.Add "removalDetails", RemovalObject(ws.Cells(r, COL_SUP_APPROVED_BY).Value, Null, ws.Cells(r, COL_SUP_REASON).Value, ws.Cells(r, COL_SUP_NUMBER_REMOVED).Value)
            d("additionalNotes") = SafeOrNull(ws.Cells(r, COL_SUP_ADDITIONAL_NOTES).Value)
            d.Add "extensions", WorkbookRowExtension(SH_SUPPLIES, r)
            coll.Add d
        End If
    Next r

    Set ExportSupplies = coll
End Function

Private Sub ImportTransactions(txs As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LEDGER)

    Dim tx As Variant
    Dim nextRow As Long

    For Each tx In txs
        nextRow = NextLedgerAppendRow(ws)
        WriteTransactionToLedgerRow ws, nextRow, tx
    Next tx
End Sub


Private Sub WriteTransactionToLedgerRow(ws As Worksheet, ByVal r As Long, tx As Variant)
    Dim wbk As Object
    Dim lines As Object
    Dim line As Variant
    Dim firstFundName As String
    Dim txBudgetName As String

    Set wbk = Nothing
    If HasWorkbookExtension(tx) Then Set wbk = tx("extensions")("workbook")

    ws.Cells(r, COL_LEDGER_TXN_DATE).Value = ParseIsoDate(ValueOrFallback(tx, "transactionDate", Null))
    ws.Cells(r, COL_LEDGER_REF).Value = ValueOrFallback(tx, "reference", "")
    ws.Cells(r, COL_LEDGER_DETAILS).Value = ValueOrFallback(tx, "description", "")

    If Not wbk Is Nothing Then
        ws.Cells(r, COL_LEDGER_DATE_SHOWS).Value = ParseIsoDate(ValueOrFallback(wbk, "dateShowsOnStatement", Null))
        ws.Cells(r, COL_LEDGER_INCOMING).Value = ParseIsoDate(ValueOrFallback(wbk, "incomingCheckOrTransferDate", Null))
        ws.Cells(r, COL_LEDGER_NAME).Value = ValueOrFallback(wbk, "personOrBusinessName", "")
        ws.Cells(r, COL_LEDGER_BANK_ACCOUNT).Value = ValueOrFallback(wbk, "bankAccount", "")
        ws.Cells(r, COL_LEDGER_AFFECTS_BANK).Value = ValueOrFallback(wbk, "affectsBank", "")
        ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value = ValueOrFallback(wbk, "budgetCategory", "")
        ws.Cells(r, COL_LEDGER_AFFECTS_BUDGET).Value = ValueOrFallback(wbk, "affectsBudget", "")
        ws.Cells(r, COL_LEDGER_FUND).Value = ValueOrFallback(wbk, "fund", "")
        ws.Cells(r, COL_LEDGER_MERCHANT).Value = ValueOrFallback(wbk, "merchant", "")
    Else
        ws.Cells(r, COL_LEDGER_AFFECTS_BANK).Value = DenormTimingValue(ValueOrFallback(tx, "bankTiming", "NONE"))
        ws.Cells(r, COL_LEDGER_AFFECTS_BUDGET).Value = DenormTimingValue(ValueOrFallback(tx, "budgetTiming", "NONE"))

        txBudgetName = LookupBudgetName(ValueOrFallback(tx, "budgetId", ""))
        If Len(txBudgetName) > 0 Then ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value = txBudgetName

        If ExistsInDict(tx, "lines") Then
            Set lines = tx("lines")
            firstFundName = FirstFundNameFromLines(lines)
            If Len(firstFundName) > 0 Then ws.Cells(r, COL_LEDGER_FUND).Value = firstFundName
        End If
    End If

    If ExistsInDict(tx, "lines") Then
        Set lines = tx("lines")
        Dim i As Long
        i = 0
        For Each line In lines
            If i > UBound(ledgerSplitCols) Then Exit For
            WriteOneSplit ws, r, ledgerSplitCols(i), line
            i = i + 1
        Next line
    End If
End Sub


Private Sub WriteOneSplit(ws As Worksheet, ByVal r As Long, grp As Variant, line As Variant)
    Dim wbk As Object
    Dim incomeCat As String
    Dim expenseCat As String
    Dim amt As Double
    Dim usedFor As Variant
    Dim itemNumber As Variant
    Dim qty As Variant

    Set wbk = Nothing
    If HasWorkbookExtension(line) Then Set wbk = line("extensions")("workbook")

    If Not wbk Is Nothing Then
        incomeCat = SafeText(ValueOrFallback(wbk, "incomeCategory", ""))
        expenseCat = SafeText(ValueOrFallback(wbk, "expenseCategory", ""))
        amt = ParseJsonNumber(ValueOrFallback(wbk, "amount", "0.00"))
        usedFor = ValueOrFallback(wbk, "usedFor", ValueOrFallback(line, "usedFor", ""))
        itemNumber = ValueOrFallback(wbk, "itemNumber", ValueOrFallback(line, "itemNumber", ""))
        qty = ValueOrFallback(wbk, "quantity", ValueOrFallback(line, "quantity", ""))
    Else
        If ParseJsonNumber(ValueOrFallback(line, "debit", "0")) > 0 Then
            expenseCat = SafeText(ValueOrFallback(line, "accountId", ""))
            incomeCat = ""
            amt = ParseJsonNumber(ValueOrFallback(line, "debit", "0"))
        Else
            incomeCat = SafeText(ValueOrFallback(line, "accountId", ""))
            expenseCat = ""
            amt = ParseJsonNumber(ValueOrFallback(line, "credit", "0"))
        End If
        usedFor = ValueOrFallback(line, "usedFor", "")
        itemNumber = ValueOrFallback(line, "itemNumber", "")
        qty = ValueOrFallback(line, "quantity", "")
    End If

    ws.Cells(r, grp(1)).Value = amt
    ws.Cells(r, grp(2)).Value = incomeCat
    ws.Cells(r, grp(3)).Value = expenseCat
    ws.Cells(r, grp(4)).Value = usedFor
    ws.Cells(r, grp(5)).Value = itemNumber
    ws.Cells(r, grp(6)).Value = qty
End Sub


Private Sub ImportOutstandingItems(items As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_OUTSTANDING)

    Dim item As Variant
    Dim r As Long

    For Each item In items
        r = NextOutstandingAppendRow(ws)
        ws.Cells(r, COL_OUT_DATE_SENT).Value = ParseIsoDate(ValueOrFallback(item, "dateSentOrReceived", Null))
        ws.Cells(r, COL_OUT_INCOMING_DATE).Value = ParseIsoDate(ValueOrFallback(item, "incomingCheckOrTransferDate", Null))
        ws.Cells(r, COL_OUT_TRANSFER_OR_CHECK).Value = ValueOrFallback(item, "transferIdOrCheckNumber", "")
        ws.Cells(r, COL_OUT_DATE_SHOWS).Value = ParseIsoDate(ValueOrFallback(item, "dateShowsOnStatement", Null))
        ws.Cells(r, COL_OUT_NAME).Value = ValueOrFallback(item, "personOrBusinessName", "")
        ws.Cells(r, COL_OUT_DETAILS).Value = ValueOrFallback(item, "detailsNotes", "")
        ws.Cells(r, COL_OUT_MERCHANT).Value = ValueOrFallback(item, "fromToCardMerchant", "")
        ws.Cells(r, COL_OUT_ACCOUNT).Value = ValueOrFallback(item, "accountForPaymentOrDeposit", "")
        ws.Cells(r, COL_OUT_AMOUNT).Value = ParseJsonNumber(ValueOrFallback(item, "amount", "0"))
        ws.Cells(r, COL_OUT_DATE_REVERSED).Value = ParseIsoDate(ValueOrFallback(item, "dateReversed", Null))
        ws.Cells(r, COL_OUT_REASON_APPROVAL).Value = ValueOrFallback(item, "reversalReasonAndApproval", "")
    Next item
End Sub


Private Sub ImportAssets(items As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_ASSETS)

    Dim item As Variant
    Dim r As Long

    For Each item In items
        r = NextAssetAppendRow(ws)
        ws.Cells(r, COL_ASSET_ITEMNUM).Value = ValueOrFallback(item, "itemNumber", "")
        ws.Cells(r, COL_ASSET_DATE_ACQ).Value = ParseIsoDate(ValueOrFallback(item, "dateAcquired", Null))
        ws.Cells(r, COL_ASSET_DESC).Value = ValueOrFallback(item, "description", "")
        ws.Cells(r, COL_ASSET_ITEM_COUNT).Value = ValueOrFallback(item, "itemCount", "")
        ws.Cells(r, COL_ASSET_TOTAL_VALUE).Value = ValueOrFallback(item, "approxValueTotal", "")
        ws.Cells(r, COL_ASSET_TOTAL_LOT_COUNT).Value = ValueOrFallback(item, "lotItemCount", "")
        ws.Cells(r, COL_ASSET_TOTAL_PAID).Value = ValueOrFallback(item, "lotPaidTotal", "")
        ws.Cells(r, COL_ASSET_PER_ITEM).Value = ValueOrFallback(item, "valuePerItem", "")
        ws.Cells(r, COL_ASSET_ITEM_TYPE).Value = ValueOrFallback(item, "itemType", "")
        ws.Cells(r, COL_ASSET_USED_FOR).Value = ValueOrFallback(item, "usedFor", "")

        If ExistsInDict(item, "currentGuardian") Then
            ws.Cells(r, COL_ASSET_GUARDIAN_NAME).Value = ValueOrFallback(item("currentGuardian"), "legalName", "")
            ws.Cells(r, COL_ASSET_GUARDIAN_EMAIL).Value = ValueOrFallback(item("currentGuardian"), "email", "")
            ws.Cells(r, COL_ASSET_GUARDIAN_PHONE).Value = ValueOrFallback(item("currentGuardian"), "phone", "")
        End If

        If ExistsInDict(item, "guardianshipDetails") Then
            ws.Cells(r, COL_ASSET_DATE_AS_OF).Value = ParseIsoDate(ValueOrFallback(item("guardianshipDetails"), "dateAsOf", Null))
            ws.Cells(r, COL_ASSET_CONFIRMED).Value = ValueOrFallback(item("guardianshipDetails"), "confirmed", "")
            ws.Cells(r, COL_ASSET_NOTES).Value = ValueOrFallback(item("guardianshipDetails"), "notes", "")
        End If

        If ExistsInDict(item, "removalDetails") Then
            ws.Cells(r, COL_ASSET_APPROVED_BY).Value = ValueOrFallback(item("removalDetails"), "approvedBy", "")
            ws.Cells(r, COL_ASSET_DATE_REMOVED).Value = ParseIsoDate(ValueOrFallback(item("removalDetails"), "approvalDate", Null))
            ws.Cells(r, COL_ASSET_REASON).Value = ValueOrFallback(item("removalDetails"), "reason", "")
            ws.Cells(r, COL_ASSET_NUM_REMOVED).Value = ValueOrFallback(item("removalDetails"), "numberRemoved", "")
        End If
    Next item
End Sub


Private Sub ImportSupplies(items As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SUPPLIES)

    Dim item As Variant
    Dim r As Long

    For Each item In items
        r = NextSupplyAppendRow(ws)
        ws.Cells(r, COL_SUP_ITEMNUM).Value = ValueOrFallback(item, "itemNumber", "")
        ws.Cells(r, COL_SUP_DATE_ACQ).Value = ParseIsoDate(ValueOrFallback(item, "dateAcquired", Null))
        ws.Cells(r, COL_SUP_DESC).Value = ValueOrFallback(item, "description", "")
        ws.Cells(r, COL_SUP_COUNT).Value = ValueOrFallback(item, "count", "")
        ws.Cells(r, COL_SUP_TOTAL_VALUE).Value = ValueOrFallback(item, "approxValueTotal", "")
        ws.Cells(r, COL_SUP_PER_ITEM).Value = ValueOrFallback(item, "valuePerItem", "")

        If ExistsInDict(item, "guardian") Then
            ws.Cells(r, COL_SUP_GUARDIAN_NAME).Value = ValueOrFallback(item("guardian"), "legalName", "")
            ws.Cells(r, COL_SUP_GUARDIAN_EMAIL).Value = ValueOrFallback(item("guardian"), "email", "")
            ws.Cells(r, COL_SUP_GUARDIAN_PHONE).Value = ValueOrFallback(item("guardian"), "phone", "")
        End If

        If ExistsInDict(item, "guardianshipDetails") Then
            ws.Cells(r, COL_SUP_DATE_AS_OF).Value = ParseIsoDate(ValueOrFallback(item("guardianshipDetails"), "dateAsOf", Null))
            ws.Cells(r, COL_SUP_LAST_CONFIRMED).Value = ParseIsoDate(ValueOrFallback(item("guardianshipDetails"), "lastConfirmed", Null))
            ws.Cells(r, COL_SUP_RETURNED).Value = ValueOrFallback(item("guardianshipDetails"), "returned", "")
            ws.Cells(r, COL_SUP_NOTES).Value = ValueOrFallback(item("guardianshipDetails"), "notes", "")
        End If

        If ExistsInDict(item, "removalDetails") Then
            ws.Cells(r, COL_SUP_APPROVED_BY).Value = ValueOrFallback(item("removalDetails"), "approvedBy", "")
            ws.Cells(r, COL_SUP_REASON).Value = ValueOrFallback(item("removalDetails"), "reason", "")
            ws.Cells(r, COL_SUP_NUMBER_REMOVED).Value = ValueOrFallback(item("removalDetails"), "numberRemoved", "")
        End If

        ws.Cells(r, COL_SUP_ADDITIONAL_NOTES).Value = ValueOrFallback(item, "additionalNotes", "")
    Next item
End Sub


Private Sub InitSplitCols()
    ledgerSplitCols = Array( _
        Array("AG", "AH", "AI", "AJ", "AK", "AL", "AM"), _
        Array("BG", "BH", "BI", "BJ", "BK", "BL", "BM"), _
        Array("CG", "CH", "CI", "CJ", "CK", "CL", "CM"), _
        Array("DG", "DH", "DI", "DJ", "DK", "DL", "DM") _
    )
End Sub

Private Function NewJsonArray() As Collection
    Set NewJsonArray = New Collection
End Function


Private Function FindLastInterestingLedgerRow(ws As Worksheet) As Long
    Dim maxR As Long
    maxR = ROW_LEDGER_FIRST
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_LEDGER_TXN_DATE).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_LEDGER_DATE_SHOWS).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_LEDGER_REF).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_LEDGER_INCOMING).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_LEDGER_NAME).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_LEDGER_DETAILS).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "AH").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "AI").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "AJ").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "BH").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "BI").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "BJ").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "CH").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "CI").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "CJ").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "DH").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "DI").End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, "DJ").End(xlUp).Row)
    FindLastInterestingLedgerRow = maxR
End Function


Private Function IsLedgerRowUsed(ws As Worksheet, ByVal r As Long) As Boolean
    If RowHasAnyValue(ws, r, Array(COL_LEDGER_TXN_DATE, COL_LEDGER_DATE_SHOWS, COL_LEDGER_REF, COL_LEDGER_INCOMING, COL_LEDGER_NAME, COL_LEDGER_DETAILS)) Then
        IsLedgerRowUsed = True
        Exit Function
    End If

    If HasSplitData(ws, r, ledgerSplitCols(0)) _
       Or HasSplitData(ws, r, ledgerSplitCols(1)) _
       Or HasSplitData(ws, r, ledgerSplitCols(2)) _
       Or HasSplitData(ws, r, ledgerSplitCols(3)) Then
        IsLedgerRowUsed = True
        Exit Function
    End If

    IsLedgerRowUsed = False
End Function


Private Function HasSplitData(ws As Worksheet, ByVal r As Long, grp As Variant) As Boolean
    If Abs(ValZero(ws.Cells(r, grp(1)).Value)) > 0 Then
        HasSplitData = True
    ElseIf Len(SafeText(ws.Cells(r, grp(2)).Value)) > 0 Then
        HasSplitData = True
    ElseIf Len(SafeText(ws.Cells(r, grp(3)).Value)) > 0 Then
        HasSplitData = True
    Else
        HasSplitData = False
    End If
End Function

Private Function FindLastUsedByAnyValue(ws As Worksheet, cols As Variant) As Long
    Dim i As Long
    Dim col As Variant
    Dim m As Long

    m = 1
    For Each col In cols
        i = ws.Cells(ws.Rows.Count, CStr(col)).End(xlUp).Row
        If i > m Then m = i
    Next col

    FindLastUsedByAnyValue = m
End Function

Private Function RowHasAnyValue(ws As Worksheet, ByVal r As Long, cols As Variant) As Boolean
    Dim col As Variant
    For Each col In cols
        If Len(Trim$(CStr(Nz(ws.Cells(r, CStr(col)).Value, "")))) > 0 Then
            RowHasAnyValue = True
            Exit Function
        End If
    Next col
    RowHasAnyValue = False
End Function


Private Function NextLedgerAppendRow(ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long

    lastRow = FindLastInterestingLedgerRow(ws)

    For r = ROW_LEDGER_FIRST To lastRow
        If Not IsLedgerRowUsed(ws, r) Then
            NextLedgerAppendRow = r
            Exit Function
        End If
    Next r

    NextLedgerAppendRow = lastRow + 1
End Function

Private Function NextAppendRow(ws As Worksheet, ByVal firstRow As Long, cols As Variant) As Long
    Dim r As Long
    Dim lastRow As Long

    lastRow = Application.Max(firstRow, FindLastUsedByAnyValue(ws, cols))

    For r = firstRow To lastRow
        If Not RowHasAnyValue(ws, r, cols) Then
            NextAppendRow = r
            Exit Function
        End If
    Next r

    NextAppendRow = lastRow + 1
End Function


Private Sub AddSimpleAccount(coll As Collection, seen As Object, ByVal nameOrId As String)
    If Len(Trim$(nameOrId)) = 0 Then Exit Sub

    Dim key As String
    key = UCase$(Trim$(nameOrId))
    If seen.Exists(key) Then Exit Sub
    seen.Add key, True

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Number") = nameOrId
    d("Name") = nameOrId
    d("Type") = GuessAccountType(nameOrId)
    d("Parent") = Null
    d("IncreaseSide") = GuessIncreaseSide(d("Type"))
    d("OpeningBalance") = "0.00"
    d.Add "SupplementalKinds", NewJsonArray()
    d("accountId") = nameOrId
    coll.Add d
End Sub

Private Function GuessAccountType(ByVal accountName As String) As String
    Dim u As String
    u = UCase$(accountName)

    If InStr(u, "CHECK") > 0 Or InStr(u, "BANK") > 0 Or InStr(u, "CASH") > 0 Or InStr(u, "ASSET") > 0 Then
        GuessAccountType = "ASSET"
    ElseIf InStr(u, "LIAB") > 0 Or InStr(u, "PAYABLE") > 0 Or InStr(u, "DEFERRED") > 0 Then
        GuessAccountType = "LIABILITY"
    ElseIf InStr(u, "REVENUE") > 0 Or InStr(u, "INCOME") > 0 Or InStr(u, "DONATION") > 0 Then
        GuessAccountType = "REVENUE"
    Else
        GuessAccountType = "EXPENSE"
    End If
End Function

Private Function GuessIncreaseSide(ByVal acctType As String) As String
    Select Case UCase$(acctType)
        Case "ASSET", "EXPENSE"
            GuessIncreaseSide = "DEBIT"
        Case Else
            GuessIncreaseSide = "CREDIT"
    End Select
End Function


Private Function GuessOutstandingKind(ws As Worksheet, ByVal r As Long) As String
    Dim ref As String
    Dim amt As Double

    ref = UCase$(SafeText(ws.Cells(r, COL_OUT_TRANSFER_OR_CHECK).Value))
    amt = CDbl(ValZero(ws.Cells(r, COL_OUT_AMOUNT).Value))

    If InStr(ref, "TR") > 0 Or InStr(ref, "XFER") > 0 Or InStr(ref, "TRANSFER") > 0 Then
        GuessOutstandingKind = "TRANSFER"
    ElseIf Len(ref) > 0 Then
        GuessOutstandingKind = "CHECK"
    ElseIf amt >= 0 Then
        GuessOutstandingKind = "DEPOSIT"
    Else
        GuessOutstandingKind = "TRANSFER"
    End If
End Function

Private Function GuessOutstandingStatus(ws As Worksheet, ByVal r As Long) As String
    If Len(SafeText(ws.Cells(r, COL_OUT_DATE_REVERSED).Value)) > 0 Then
        GuessOutstandingStatus = "REVERSED"
    ElseIf Len(SafeText(ws.Cells(r, COL_OUT_DATE_SHOWS).Value)) > 0 Then
        GuessOutstandingStatus = "CLEARED"
    Else
        GuessOutstandingStatus = "OUTSTANDING"
    End If
End Function


Private Function WorkbookRowExtension(ByVal sheetName As String, ByVal rowNum As Long) As Object
    Dim ext As Object
    Dim wbk As Object

    Set ext = CreateObject("Scripting.Dictionary")
    Set wbk = CreateObject("Scripting.Dictionary")
    wbk("sheet") = sheetName
    wbk("row") = rowNum
    ext.Add "workbook", wbk

    Set WorkbookRowExtension = ext
End Function

Private Function GuardianObject(nm As Variant, em As Variant, ph As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("legalName") = SafeOrNull(nm)
    d("email") = SafeOrNull(em)
    d("phone") = SafeOrNull(ph)
    Set GuardianObject = d
End Function

Private Function GuardianshipObject(dt As Variant, confirmed As Variant, notes As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("dateAsOf") = ISODateOrNull(dt)
    d("confirmed") = BoolOrNull(confirmed)
    d("notes") = SafeOrNull(notes)
    Set GuardianshipObject = d
End Function

Private Function RemovalObject(approvedBy As Variant, approvalDate As Variant, reason As Variant, numberRemoved As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("approvedBy") = SafeOrNull(approvedBy)
    d("approvalDate") = ISODateOrNull(approvalDate)
    d("reason") = SafeOrNull(reason)
    d("numberRemoved") = NullOrNumber(numberRemoved)
    Set RemovalObject = d
End Function

Private Function HasWorkbookExtension(obj As Variant) As Boolean
    On Error GoTo Nope
    HasWorkbookExtension = ExistsInDict(obj, "extensions") And ExistsInDict(obj("extensions"), "workbook")
    Exit Function
Nope:
    HasWorkbookExtension = False
End Function

Private Function ExistsInDict(obj As Variant, ByVal key As String) As Boolean
    On Error Resume Next
    ExistsInDict = obj.Exists(key)
    On Error GoTo 0
End Function

Private Function ValueOrFallback(obj As Variant, ByVal key As String, fallback As Variant) As Variant
    If ExistsInDict(obj, key) Then
        ValueOrFallback = obj(key)
    Else
        ValueOrFallback = fallback
    End If
End Function

Private Function SafeText(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeText = ""
    Else
        SafeText = Trim$(CStr(v))
    End If
End Function

Private Function SafeOrNull(v As Variant) As Variant
    Dim s As String
    s = SafeText(v)

    If Len(s) = 0 Then
        SafeOrNull = Null
    Else
        SafeOrNull = s
    End If
End Function

Private Function NullIfEmpty(ByVal s As String) As Variant
    If Len(Trim$(s)) = 0 Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = s
    End If
End Function

Private Function Nz(v As Variant, Optional fallback As Variant = "") As Variant
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Nz = fallback
    Else
        Nz = v
    End If
End Function

Private Function ValZero(v As Variant) As Double
    ValZero = CDbl(Val(Replace(SafeText(v), ",", "")))
End Function

Private Function ParseJsonNumber(ByVal v As Variant) As Double
    ParseJsonNumber = CDbl(Val(Replace(SafeText(v), ",", "")))
End Function

Private Function NullOrNumber(v As Variant) As Variant
    If Len(SafeText(v)) = 0 Then
        NullOrNumber = Null
    Else
        NullOrNumber = v
    End If
End Function

Private Function BoolOrNull(v As Variant) As Variant
    Dim s As String
    s = UCase$(SafeText(v))

    If Len(s) = 0 Then
        BoolOrNull = Null
    ElseIf s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Then
        BoolOrNull = True
    ElseIf s = "FALSE" Or s = "NO" Or s = "N" Or s = "0" Then
        BoolOrNull = False
    ElseIf VarType(v) = vbBoolean Then
        BoolOrNull = CBool(v)
    Else
        BoolOrNull = Null
    End If
End Function

Private Function AmountOrNull(v As Variant) As Variant
    If Len(SafeText(v)) = 0 Then
        AmountOrNull = Null
    Else
        AmountOrNull = FormatAmount(CDbl(ValZero(v)))
    End If
End Function

Private Function FormatAmount(ByVal n As Double) As String
    FormatAmount = Format$(n, "0.00")
End Function

Private Function FormatAmountAbs(v As Variant) As String
    FormatAmountAbs = Format$(Abs(CDbl(ValZero(v))), "0.00")
End Function

Private Function ISODateOrNull(v As Variant) As Variant
    If IsDate(v) Then
        ISODateOrNull = Format$(CDate(v), "yyyy-mm-dd")
    ElseIf Len(SafeText(v)) = 0 Then
        ISODateOrNull = Null
    Else
        ISODateOrNull = SafeText(v)
    End If
End Function


Private Function ParseIsoDate(v As Variant) As Variant
    Dim s As String
    Dim y As Integer
    Dim m As Integer
    Dim d As Integer

    s = SafeText(v)

    If Len(s) = 0 Then
        ParseIsoDate = vbNullString
        Exit Function
    End If

    s = Left$(s, 10)

    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        On Error GoTo BadDate
        y = CInt(Left$(s, 4))
        m = CInt(Mid$(s, 6, 2))
        d = CInt(Right$(s, 2))
        ParseIsoDate = DateSerial(y, m, d)
        Exit Function
    End If

BadDate:
    ParseIsoDate = vbNullString
End Function

Private Function NormalizeId(ByVal prefix As String, ByVal raw As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim body As String

    s = LCase$(Trim$(raw))
    body = ""

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "a" To "z", "0" To "9"
                body = body & ch
            Case Else
                body = body & "-"
        End Select
    Next i

    Do While InStr(body, "--") > 0
        body = Replace(body, "--", "-")
    Loop

    Do While Left$(body, 1) = "-"
        body = Mid$(body, 2)
        If Len(body) = 0 Then Exit Do
    Loop

    Do While Right$(body, 1) = "-"
        body = Left$(body, Len(body) - 1)
        If Len(body) = 0 Then Exit Do
    Loop

    NormalizeId = prefix & body
End Function


Private Function YearStartFromSummary(ws As Worksheet) As String
    YearStartFromSummary = Format$(DateSerial(SummaryYear(ws), 1, 1), "yyyy-mm-dd")
End Function


Private Function YearEndFromSummary(ws As Worksheet) As String
    If IsDate(ws.Range(CELL_SUM_FY_END).Value) Then
        YearEndFromSummary = Format$(CDate(ws.Range(CELL_SUM_FY_END).Value), "yyyy-mm-dd")
    Else
        YearEndFromSummary = Format$(DateSerial(SummaryYear(ws), 12, 31), "yyyy-mm-dd")
    End If
End Function


Private Function SummaryYear(ws As Worksheet) As Long
    Dim yr As Long
    yr = CLng(ValZero(ws.Range(CELL_SUM_LEDGER_YEAR).Value))
    If yr = 0 Then yr = Year(Date)
    SummaryYear = yr
End Function

Private Function DateCellOrFallback(v As Variant, ByVal fallbackIso As String) As String
    If IsDate(v) Then
        DateCellOrFallback = Format$(CDate(v), "yyyy-mm-dd")
    Else
        DateCellOrFallback = fallbackIso
    End If
End Function

Private Function FormatDateTimeOffset(ByVal dt As Date) As String
    FormatDateTimeOffset = Format$(dt, "yyyy-mm-dd\Thh:nn:ss") & "Z"
End Function

Private Function WorkbookLinkObject(ByVal sheetKey As String, ByVal rowNum As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("sheetKey") = sheetKey
    d("ledgerRowIndex") = rowNum
    Set WorkbookLinkObject = d
End Function

Private Function TimingValueFromWorkbook(v As Variant) As String
    Dim s As String
    s = UCase$(SafeText(v))

    Select Case s
        Case "NOW"
            TimingValueFromWorkbook = "NOW"
        Case "PREVIOUSLY", "PREVIOUS"
            TimingValueFromWorkbook = "PREVIOUSLY"
        Case "LATER"
            TimingValueFromWorkbook = "LATER"
        Case Else
            TimingValueFromWorkbook = "NONE"
    End Select
End Function

Private Function DenormTimingValue(v As Variant) As String
    Select Case UCase$(SafeText(v))
        Case "NOW"
            DenormTimingValue = "Now"
        Case "PREVIOUSLY"
            DenormTimingValue = "Previously"
        Case "LATER"
            DenormTimingValue = "Later"
        Case Else
            DenormTimingValue = ""
    End Select
End Function

Private Function BudgetIdFromFields(ByVal budgetName As String, ByVal fundName As String) As Variant
    If Len(Trim$(budgetName)) = 0 Then
        BudgetIdFromFields = Null
    ElseIf Len(Trim$(fundName)) = 0 Then
        BudgetIdFromFields = NormalizeId("budget-", budgetName)
    Else
        BudgetIdFromFields = NormalizeId("budget-", fundName & "-" & budgetName)
    End If
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        MaxLong = a
    Else
        MaxLong = b
    End If
End Function

Private Function FindLastInterestingOutstandingRow(ws As Worksheet) As Long
    Dim maxR As Long
    maxR = ROW_OUT_FIRST
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_DATE_SENT).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_INCOMING_DATE).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_TRANSFER_OR_CHECK).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_DATE_SHOWS).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_NAME).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_DETAILS).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_AMOUNT).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_DATE_REVERSED).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_OUT_REASON_APPROVAL).End(xlUp).Row)
    FindLastInterestingOutstandingRow = maxR
End Function

Private Function IsOutstandingRowUsed(ws As Worksheet, ByVal r As Long) As Boolean
    If Abs(ValZero(ws.Cells(r, COL_OUT_AMOUNT).Value)) > 0 Then
        IsOutstandingRowUsed = True
    Else
        IsOutstandingRowUsed = RowHasAnyValue(ws, r, Array(COL_OUT_DATE_SENT, COL_OUT_INCOMING_DATE, COL_OUT_TRANSFER_OR_CHECK, COL_OUT_DATE_SHOWS, COL_OUT_NAME, COL_OUT_DETAILS, COL_OUT_DATE_REVERSED, COL_OUT_REASON_APPROVAL))
    End If
End Function

Private Function NextOutstandingAppendRow(ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long
    lastRow = FindLastInterestingOutstandingRow(ws)

    For r = ROW_OUT_FIRST To lastRow
        If Not IsOutstandingRowUsed(ws, r) Then
            NextOutstandingAppendRow = r
            Exit Function
        End If
    Next r

    NextOutstandingAppendRow = lastRow + 1
End Function

Private Function FindLastInterestingAssetRow(ws As Worksheet) As Long
    Dim maxR As Long
    maxR = ROW_ASSET_FIRST
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_DATE_ACQ).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_DESC).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_ITEM_COUNT).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_TOTAL_VALUE).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_TOTAL_LOT_COUNT).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_TOTAL_PAID).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_PER_ITEM).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_GUARDIAN_NAME).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_GUARDIAN_EMAIL).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_GUARDIAN_PHONE).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_DATE_AS_OF).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_NOTES).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_APPROVED_BY).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_DATE_REMOVED).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_ASSET_NUM_REMOVED).End(xlUp).Row)
    FindLastInterestingAssetRow = maxR
End Function

Private Function IsAssetRowUsed(ws As Worksheet, ByVal r As Long) As Boolean
    If RowHasAnyValue(ws, r, Array(COL_ASSET_DATE_ACQ, COL_ASSET_DESC, COL_ASSET_GUARDIAN_NAME, COL_ASSET_GUARDIAN_EMAIL, COL_ASSET_GUARDIAN_PHONE, COL_ASSET_DATE_AS_OF, COL_ASSET_NOTES, COL_ASSET_APPROVED_BY, COL_ASSET_DATE_REMOVED, COL_ASSET_NUM_REMOVED)) Then
        IsAssetRowUsed = True
    ElseIf Abs(ValZero(ws.Cells(r, COL_ASSET_ITEM_COUNT).Value)) > 0 _
        Or Abs(ValZero(ws.Cells(r, COL_ASSET_TOTAL_VALUE).Value)) > 0 _
        Or Abs(ValZero(ws.Cells(r, COL_ASSET_TOTAL_LOT_COUNT).Value)) > 0 _
        Or Abs(ValZero(ws.Cells(r, COL_ASSET_TOTAL_PAID).Value)) > 0 _
        Or Abs(ValZero(ws.Cells(r, COL_ASSET_PER_ITEM).Value)) > 0 Then
        IsAssetRowUsed = True
    Else
        IsAssetRowUsed = False
    End If
End Function

Private Function NextAssetAppendRow(ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long
    lastRow = FindLastInterestingAssetRow(ws)

    For r = ROW_ASSET_FIRST To lastRow
        If Not IsAssetRowUsed(ws, r) Then
            NextAssetAppendRow = r
            Exit Function
        End If
    Next r

    NextAssetAppendRow = lastRow + 1
End Function

Private Function FindLastInterestingSupplyRow(ws As Worksheet) As Long
    Dim maxR As Long
    maxR = ROW_SUPPLY_FIRST
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_DATE_ACQ).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_DESC).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_COUNT).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_TOTAL_VALUE).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_PER_ITEM).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_GUARDIAN_NAME).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_GUARDIAN_EMAIL).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_GUARDIAN_PHONE).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_DATE_AS_OF).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_LAST_CONFIRMED).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_NOTES).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_APPROVED_BY).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_NUMBER_REMOVED).End(xlUp).Row)
    maxR = MaxLong(maxR, ws.Cells(ws.Rows.Count, COL_SUP_ADDITIONAL_NOTES).End(xlUp).Row)
    FindLastInterestingSupplyRow = maxR
End Function

Private Function IsSupplyRowUsed(ws As Worksheet, ByVal r As Long) As Boolean
    If RowHasAnyValue(ws, r, Array(COL_SUP_DATE_ACQ, COL_SUP_DESC, COL_SUP_GUARDIAN_NAME, COL_SUP_GUARDIAN_EMAIL, COL_SUP_GUARDIAN_PHONE, COL_SUP_DATE_AS_OF, COL_SUP_LAST_CONFIRMED, COL_SUP_NOTES, COL_SUP_APPROVED_BY, COL_SUP_NUMBER_REMOVED, COL_SUP_ADDITIONAL_NOTES)) Then
        IsSupplyRowUsed = True
    ElseIf Abs(ValZero(ws.Cells(r, COL_SUP_COUNT).Value)) > 0 _
        Or Abs(ValZero(ws.Cells(r, COL_SUP_TOTAL_VALUE).Value)) > 0 _
        Or Abs(ValZero(ws.Cells(r, COL_SUP_PER_ITEM).Value)) > 0 Then
        IsSupplyRowUsed = True
    Else
        IsSupplyRowUsed = False
    End If
End Function

Private Function NextSupplyAppendRow(ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long
    lastRow = FindLastInterestingSupplyRow(ws)

    For r = ROW_SUPPLY_FIRST To lastRow
        If Not IsSupplyRowUsed(ws, r) Then
            NextSupplyAppendRow = r
            Exit Function
        End If
    Next r

    NextSupplyAppendRow = lastRow + 1
End Function

Private Sub ClearImportLookupMaps()
    Set mBudgetNameById = CreateObject("Scripting.Dictionary")
    Set mFundNameById = CreateObject("Scripting.Dictionary")
End Sub

Private Sub BuildImportLookupMaps(root As Object)
    Dim item As Variant

    If mBudgetNameById Is Nothing Then Set mBudgetNameById = CreateObject("Scripting.Dictionary")
    If mFundNameById Is Nothing Then Set mFundNameById = CreateObject("Scripting.Dictionary")

    If ExistsInDict(root, "budgets") Then
        For Each item In root("budgets")
            If ExistsInDict(item, "budgetId") And ExistsInDict(item, "name") Then
                mBudgetNameById(UCase$(CStr(item("budgetId")))) = CStr(item("name"))
            End If
        Next item
    End If

    If ExistsInDict(root, "funds") Then
        For Each item In root("funds")
            If ExistsInDict(item, "fundId") And ExistsInDict(item, "name") Then
                mFundNameById(UCase$(CStr(item("fundId")))) = CStr(item("name"))
            End If
        Next item
    End If
End Sub

Private Function LookupBudgetName(ByVal budgetId As String) As String
    If mBudgetNameById Is Nothing Then
        LookupBudgetName = ""
    ElseIf mBudgetNameById.Exists(UCase$(budgetId)) Then
        LookupBudgetName = CStr(mBudgetNameById(UCase$(budgetId)))
    Else
        LookupBudgetName = ""
    End If
End Function

Private Function LookupFundName(ByVal fundId As String) As String
    If mFundNameById Is Nothing Then
        LookupFundName = ""
    ElseIf mFundNameById.Exists(UCase$(fundId)) Then
        LookupFundName = CStr(mFundNameById(UCase$(fundId)))
    Else
        LookupFundName = ""
    End If
End Function

Private Function FirstFundNameFromLines(lines As Object) As String
    Dim line As Variant
    Dim fundId As String

    For Each line In lines
        fundId = SafeText(ValueOrFallback(line, "fundId", ""))
        If Len(fundId) > 0 Then
            FirstFundNameFromLines = LookupFundName(fundId)
            If Len(FirstFundNameFromLines) = 0 Then FirstFundNameFromLines = fundId
            Exit Function
        End If
    Next line
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal text As String)
    Dim ff As Integer
    ff = FreeFile

    On Error GoTo CleanFail
    Open path For Output As #ff
    Print #ff, text;
    Close #ff
    Exit Sub

CleanFail:
    On Error Resume Next
    Close #ff
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim ff As Integer
    ff = FreeFile

    On Error GoTo CleanFail
    Open path For Input As #ff
    ReadTextFile = Input$(LOF(ff), ff)
    Close #ff
    Exit Function

CleanFail:
    On Error Resume Next
    Close #ff
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function