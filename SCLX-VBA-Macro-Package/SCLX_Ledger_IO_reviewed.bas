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

Private Const SCLX_VERSION As String = "1.2"

Private Const SH_SUMMARY As String = "Summary"
Private Const SH_LEDGER As String = "Ledger"
Private Const SH_OUTSTANDING As String = "Outstanding"
Private Const SH_ASSETS As String = "Assets&Inventory"
Private Const SH_SUPPLIES As String = "Supplies"

Private Const ROW_LEDGER_FIRST As Long = 9
Private Const ROW_OUT_FIRST As Long = 14
Private Const ROW_ASSET_FIRST As Long = 11
Private Const ROW_SUPPLY_FIRST As Long = 10

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
'   0 Amount
'   1 Income Category
'   2 Expense Category
'   3 Used For
'   4 Item Num
'   5 Qty
'   6 Spare / notes / reserved workbook column
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
    path = Application.GetSaveAsFilename(InitialFileName:="ledger.sclx.json", _
                                         FileFilter:="JSON Files (*.json), *.json")
    If VarType(path) = vbBoolean Then Exit Sub

    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")

    root("format") = "SCLX"
    root("version") = SCLX_VERSION
    root("exportedAt") = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    root("organization") = ExportOrganization()
    root("reportingPeriod") = ExportReportingPeriod()
    root("chartOfAccounts") = ExportChartOfAccounts()
    root("funds") = ExportFunds()
    root("budgets") = ExportBudgets()
    root("people") = NewJsonArray()
    root("events") = NewJsonArray()
    root("documents") = NewJsonArray()
    root("transactions") = ExportTransactions()
    root("bankingItems") = NewJsonArray()
    root("outstandingItems") = ExportOutstandingItems()
    root("otherAssetItems") = NewJsonArray()
    root("assets") = ExportAssets()
    root("supplies") = ExportSupplies()
    root("bankStatementImports") = NewJsonArray()
    root("extensions") = CreateObject("Scripting.Dictionary")

    WriteTextFile CStr(path), JsonConverter.ConvertToJson(root, Whitespace:=2)
    MsgBox "SCLX export written to:" & vbCrLf & CStr(path), vbInformation
    Exit Sub

EH:
    MsgBox "ExportSCLX failed: " & Err.Description, vbCritical
End Sub

Public Sub ImportSCLX()
    On Error GoTo EH

    InitSplitCols

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

'========================
' Export helpers
'========================

Private Function ExportOrganization() As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SUMMARY)

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("organizationId") = SafeText(ws.Range("B3").Value)
    d("name") = SafeText(ws.Range("B3").Value)
    d("parentOrganization") = Null
    d("baseCurrency") = SafeText(ws.Range("H8").Value)
    d("fiscalYearStart") = YearStartFromSummary(ws)
    d("fiscalYearEnd") = YearEndFromSummary(ws)

    Set ExportOrganization = d
End Function

Private Function ExportReportingPeriod() As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SUMMARY)

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("startDate") = YearStartFromSummary(ws)
    d("endDate") = YearEndFromSummary(ws)
    d("label") = SafeText(ws.Range("H6").Value) & " Q" & SafeText(ws.Range("H7").Value)

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
    Set ExportBudgets = New Collection
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

    Set tx = CreateObject("Scripting.Dictionary")
    tx("transactionId") = "ledger-row-" & CStr(r)
    tx("transactionDate") = ISODateOrNull(ws.Cells(r, COL_LEDGER_TXN_DATE).Value)
    tx("postingDate") = ISODateOrNull(ws.Cells(r, COL_LEDGER_TXN_DATE).Value)
    tx("description") = SafeText(ws.Cells(r, COL_LEDGER_DETAILS).Value)
    tx("reference") = SafeText(ws.Cells(r, COL_LEDGER_REF).Value)
    tx("status") = "POSTED"
    tx("source") = "MANUAL"

    Set ext = CreateObject("Scripting.Dictionary")
    Set wbk = CreateObject("Scripting.Dictionary")

    wbk("ledgerRow") = r
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

    ext("workbook") = wbk
    tx("extensions") = ext
    tx("lines") = ExportLedgerSplitLines(ws, r)

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

    amt = CDbl(ValZero(ws.Cells(r, grp(0)).Value))
    incomeCat = SafeText(ws.Cells(r, grp(1)).Value)
    expenseCat = SafeText(ws.Cells(r, grp(2)).Value)

    Set d = CreateObject("Scripting.Dictionary")
    d("lineId") = "ledger-row-" & r & "-ln-" & splitIndex

    If Len(expenseCat) > 0 Then
        acct = expenseCat
    Else
        acct = incomeCat
    End If
    If Len(acct) = 0 Then acct = "UNMAPPED"

    d("accountId") = acct

    If Len(expenseCat) > 0 Then
        d("debit") = FormatAmount(Abs(amt))
        d("credit") = FormatAmount(0)
    Else
        d("debit") = FormatAmount(0)
        d("credit") = FormatAmount(Abs(amt))
    End If

    If Len(SafeText(ws.Cells(r, COL_LEDGER_FUND).Value)) > 0 Then
        d("fundId") = NormalizeId("fund-", SafeText(ws.Cells(r, COL_LEDGER_FUND).Value))
    Else
        d("fundId") = Null
    End If

    d("budgetId") = SafeOrNull(ws.Cells(r, COL_LEDGER_BUDGET_CATEGORY).Value)
    d("personId") = SafeOrNull(ws.Cells(r, COL_LEDGER_NAME).Value)
    d("eventId") = Null
    d("documentId") = Null
    d("memo") = SafeText(ws.Cells(r, COL_LEDGER_DETAILS).Value)

    Set ext = CreateObject("Scripting.Dictionary")
    Set wbk = CreateObject("Scripting.Dictionary")

    wbk("splitIndex") = splitIndex
    wbk("amount") = FormatAmount(Abs(amt))
    wbk("incomeCategory") = NullIfEmpty(incomeCat)
    wbk("expenseCategory") = NullIfEmpty(expenseCat)
    wbk("usedFor") = SafeOrNull(ws.Cells(r, grp(3)).Value)
    wbk("itemNumber") = SafeOrNull(ws.Cells(r, grp(4)).Value)
    wbk("quantity") = SafeOrNull(ws.Cells(r, grp(5)).Value)
    wbk("reserved") = SafeOrNull(ws.Cells(r, grp(6)).Value)

    ext("workbook") = wbk
    d("extensions") = ext

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

    lastRow = FindLastUsedByAnyValue(ws, Array(COL_OUT_DATE_SENT, COL_OUT_NAME, COL_OUT_AMOUNT))

    For r = ROW_OUT_FIRST To lastRow
        If RowHasAnyValue(ws, r, Array(COL_OUT_DATE_SENT, COL_OUT_NAME, COL_OUT_AMOUNT, COL_OUT_TRANSFER_OR_CHECK)) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("outstandingItemId") = "outstanding-row-" & r
            d("kind") = GuessOutstandingKind(ws, r)
            d("ledgerLink") = Null
            d("amount") = FormatAmountAbs(ws.Cells(r, COL_OUT_AMOUNT).Value)

            Set ext = CreateObject("Scripting.Dictionary")
            Set wbk = CreateObject("Scripting.Dictionary")
            wbk("sheet") = SH_OUTSTANDING
            wbk("row") = r
            ext("workbook") = wbk
            d("extensions") = ext

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

    lastRow = FindLastUsedByAnyValue(ws, Array(COL_ASSET_DATE_ACQ, COL_ASSET_DESC, COL_ASSET_GUARDIAN_NAME))

    For r = ROW_ASSET_FIRST To lastRow
        If RowHasAnyValue(ws, r, Array(COL_ASSET_DATE_ACQ, COL_ASSET_DESC, COL_ASSET_ITEM_TYPE, COL_ASSET_GUARDIAN_NAME)) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("assetId") = "asset-row-" & r
            d("itemNumber") = SafeOrNull(ws.Cells(r, COL_ASSET_ITEMNUM).Value)
            d("dateAcquired") = ISODateOrNull(ws.Cells(r, COL_ASSET_DATE_ACQ).Value)
            d("description") = SafeOrNull(ws.Cells(r, COL_ASSET_DESC).Value)
            d("itemCount") = NullOrNumber(ws.Cells(r, COL_ASSET_ITEM_COUNT).Value)
            d("approxValueTotal") = AmountOrNull(ws.Cells(r, COL_ASSET_TOTAL_VALUE).Value)
            d("valuePerItem") = AmountOrNull(ws.Cells(r, COL_ASSET_PER_ITEM).Value)
            d("itemType") = SafeOrNull(ws.Cells(r, COL_ASSET_ITEM_TYPE).Value)
            d("usedFor") = SafeOrNull(ws.Cells(r, COL_ASSET_USED_FOR).Value)
            d("lotPaidTotal") = AmountOrNull(ws.Cells(r, COL_ASSET_TOTAL_PAID).Value)
            d("lotItemCount") = NullOrNumber(ws.Cells(r, COL_ASSET_TOTAL_LOT_COUNT).Value)
            d("currentGuardian") = GuardianObject(ws.Cells(r, COL_ASSET_GUARDIAN_NAME).Value, ws.Cells(r, COL_ASSET_GUARDIAN_EMAIL).Value, ws.Cells(r, COL_ASSET_GUARDIAN_PHONE).Value)
            d("guardianshipDetails") = GuardianshipObject(ws.Cells(r, COL_ASSET_DATE_AS_OF).Value, ws.Cells(r, COL_ASSET_CONFIRMED).Value, ws.Cells(r, COL_ASSET_NOTES).Value)
            d("removalDetails") = RemovalObject(ws.Cells(r, COL_ASSET_APPROVED_BY).Value, ws.Cells(r, COL_ASSET_DATE_REMOVED).Value, ws.Cells(r, COL_ASSET_REASON).Value, ws.Cells(r, COL_ASSET_NUM_REMOVED).Value)
            d("extensions") = WorkbookRowExtension(SH_ASSETS, r)
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

    lastRow = FindLastUsedByAnyValue(ws, Array(COL_SUP_DATE_ACQ, COL_SUP_DESC, COL_SUP_GUARDIAN_NAME))

    For r = ROW_SUPPLY_FIRST To lastRow
        If RowHasAnyValue(ws, r, Array(COL_SUP_DATE_ACQ, COL_SUP_DESC, COL_SUP_GUARDIAN_NAME, COL_SUP_REASON)) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("supplyId") = "supply-row-" & r
            d("itemNumber") = SafeOrNull(ws.Cells(r, COL_SUP_ITEMNUM).Value)
            d("dateAcquired") = ISODateOrNull(ws.Cells(r, COL_SUP_DATE_ACQ).Value)
            d("description") = SafeOrNull(ws.Cells(r, COL_SUP_DESC).Value)
            d("count") = NullOrNumber(ws.Cells(r, COL_SUP_COUNT).Value)
            d("approxValueTotal") = AmountOrNull(ws.Cells(r, COL_SUP_TOTAL_VALUE).Value)
            d("valuePerItem") = AmountOrNull(ws.Cells(r, COL_SUP_PER_ITEM).Value)
            d("guardian") = GuardianObject(ws.Cells(r, COL_SUP_GUARDIAN_NAME).Value, ws.Cells(r, COL_SUP_GUARDIAN_EMAIL).Value, ws.Cells(r, COL_SUP_GUARDIAN_PHONE).Value)

            Set gd = CreateObject("Scripting.Dictionary")
            gd("dateAsOf") = ISODateOrNull(ws.Cells(r, COL_SUP_DATE_AS_OF).Value)
            gd("lastConfirmed") = ISODateOrNull(ws.Cells(r, COL_SUP_LAST_CONFIRMED).Value)
            gd("returned") = BoolOrNull(ws.Cells(r, COL_SUP_RETURNED).Value)
            gd("notes") = SafeOrNull(ws.Cells(r, COL_SUP_NOTES).Value)
            d("guardianshipDetails") = gd

            d("removalDetails") = RemovalObject(ws.Cells(r, COL_SUP_APPROVED_BY).Value, Null, ws.Cells(r, COL_SUP_REASON).Value, ws.Cells(r, COL_SUP_NUMBER_REMOVED).Value)
            d("additionalNotes") = SafeOrNull(ws.Cells(r, COL_SUP_ADDITIONAL_NOTES).Value)
            d("extensions") = WorkbookRowExtension(SH_SUPPLIES, r)
            coll.Add d
        End If
    Next r

    Set ExportSupplies = coll
End Function

'========================
' Import helpers
'========================

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
    Set wbk = Nothing

    If HasWorkbookExtension(tx) Then
        Set wbk = tx("extensions")("workbook")
    End If

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
    End If

    Dim lines As Object
    Dim i As Long
    Dim line As Variant

    If ExistsInDict(tx, "lines") Then
        Set lines = tx("lines")
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

    Set wbk = Nothing

    If HasWorkbookExtension(line) Then
        Set wbk = line("extensions")("workbook")
    End If

    If Not wbk Is Nothing Then
        incomeCat = SafeText(ValueOrFallback(wbk, "incomeCategory", ""))
        expenseCat = SafeText(ValueOrFallback(wbk, "expenseCategory", ""))
        amt = ParseJsonNumber(ValueOrFallback(wbk, "amount", "0.00"))
        ws.Cells(r, grp(3)).Value = ValueOrFallback(wbk, "usedFor", "")
        ws.Cells(r, grp(4)).Value = ValueOrFallback(wbk, "itemNumber", "")
        ws.Cells(r, grp(5)).Value = ValueOrFallback(wbk, "quantity", "")
        ws.Cells(r, grp(6)).Value = ValueOrFallback(wbk, "reserved", "")
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
    End If

    ws.Cells(r, grp(0)).Value = amt
    ws.Cells(r, grp(1)).Value = incomeCat
    ws.Cells(r, grp(2)).Value = expenseCat
End Sub

Private Sub ImportOutstandingItems(items As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_OUTSTANDING)

    Dim item As Variant
    Dim r As Long

    For Each item In items
        r = NextAppendRow(ws, ROW_OUT_FIRST, Array(COL_OUT_DATE_SENT, COL_OUT_NAME, COL_OUT_AMOUNT))
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
        r = NextAppendRow(ws, ROW_ASSET_FIRST, Array(COL_ASSET_DATE_ACQ, COL_ASSET_DESC, COL_ASSET_GUARDIAN_NAME))
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
        r = NextAppendRow(ws, ROW_SUPPLY_FIRST, Array(COL_SUP_DATE_ACQ, COL_SUP_DESC, COL_SUP_GUARDIAN_NAME))
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

'========================
' Utility helpers
'========================

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
    maxR = ws.Cells(ws.Rows.Count, COL_LEDGER_TXN_DATE).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row > maxR Then maxR = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "BG").End(xlUp).Row > maxR Then maxR = ws.Cells(ws.Rows.Count, "BG").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "CG").End(xlUp).Row > maxR Then maxR = ws.Cells(ws.Rows.Count, "CG").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "DG").End(xlUp).Row > maxR Then maxR = ws.Cells(ws.Rows.Count, "DG").End(xlUp).Row
    If maxR < ROW_LEDGER_FIRST Then maxR = ROW_LEDGER_FIRST
    FindLastInterestingLedgerRow = maxR
End Function

Private Function IsLedgerRowUsed(ws As Worksheet, ByVal r As Long) As Boolean
    IsLedgerRowUsed = RowHasAnyValue(ws, r, Array(COL_LEDGER_TXN_DATE, COL_LEDGER_REF, COL_LEDGER_NAME, COL_LEDGER_DETAILS, "AG", "BG", "CG", "DG"))
End Function

Private Function HasSplitData(ws As Worksheet, ByVal r As Long, grp As Variant) As Boolean
    HasSplitData = RowHasAnyValue(ws, r, Array(grp(0), grp(1), grp(2), grp(3), grp(4), grp(5), grp(6)))
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
    NextLedgerAppendRow = NextAppendRow(ws, ROW_LEDGER_FIRST, Array(COL_LEDGER_TXN_DATE, COL_LEDGER_REF, COL_LEDGER_NAME, COL_LEDGER_DETAILS, "AG", "BG", "CG", "DG"))
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
    d("number") = nameOrId
    d("name") = nameOrId
    d("type") = GuessAccountType(nameOrId)
    d("parent") = Null
    d("increaseSide") = GuessIncreaseSide(d("type"))
    d("openingBalance") = "0.00"
    d("supplementalKinds") = NewJsonArray()
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
    ext("workbook") = wbk

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
    Dim yr As Long
    yr = CLng(ValZero(ws.Range("H6").Value))
    If yr = 0 Then yr = Year(Date)
    YearStartFromSummary = Format$(DateSerial(yr, 1, 1), "yyyy-mm-dd")
End Function

Private Function YearEndFromSummary(ws As Worksheet) As String
    Dim yr As Long
    yr = CLng(ValZero(ws.Range("H6").Value))
    If yr = 0 Then yr = Year(Date)
    YearEndFromSummary = Format$(DateSerial(yr, 12, 31), "yyyy-mm-dd")
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