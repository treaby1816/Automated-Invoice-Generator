Attribute VB_Name = "Module1"
Option Explicit

' ================================================================
' BUSINESS LOGIC MACROS
' ================================================================

Public Sub AddStockItem()
    On Error GoTo ErrorHandler

    Dim wsStock As Worksheet
    Dim wsInv As Worksheet
    Dim wsSet As Worksheet
    
    Set wsStock = ThisWorkbook.Sheets("StockIn")
    Set wsInv = ThisWorkbook.Sheets("Inventory")
    Set wsSet = ThisWorkbook.Sheets("Settings")

    ' Read values from the In-Sheet Form on StockIn sheet
    Dim pName As String: pName = Trim(wsStock.Range("D4").Value)
    Dim pCat As String: pCat = Trim(wsStock.Range("D5").Value)
    Dim pDesc As String: pDesc = Trim(wsStock.Range("D6").Value)
    Dim pQty As Variant: pQty = wsStock.Range("D7").Value
    Dim pCost As Variant: pCost = wsStock.Range("D8").Value
    Dim pUser As String: pUser = Trim(wsStock.Range("D9").Value)

    ' 1. Validate required fields
    If pName = "" Or pCat = "" Or IsEmpty(pQty) Or IsEmpty(pCost) Then
        MsgBox "Please fill in all required fields (*).", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' 2. Validate numeric inputs
    If Not IsNumeric(pQty) Or Not IsNumeric(pCost) Then
        MsgBox "Quantity and Unit Cost must be valid numbers.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    If CDbl(pQty) <= 0 Or CDbl(pCost) < 0 Then
        MsgBox "Quantity must be greater than 0, and Cost cannot be negative.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' 3. Check if Product exists in Inventory
    Dim matchRow As Variant
    matchRow = Application.Match(pName, wsInv.Columns("B"), 0)

    Dim pID As String

    If IsError(matchRow) Then
        ' NEW PRODUCT
        Dim nextStkNum As Long
        nextStkNum = wsSet.Range("B7").Value
        pID = "PRD-" & Format(nextStkNum, "000")
        
        Dim newRow As Long
        newRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row + 1
        
        wsInv.Cells(newRow, 1).Value = pID
        wsInv.Cells(newRow, 2).Value = pName
        wsInv.Cells(newRow, 3).Value = pCat
        wsInv.Cells(newRow, 4).Value = pDesc
        wsInv.Cells(newRow, 5).Value = CDbl(pCost)
        wsInv.Cells(newRow, 6).Value = CDbl(pQty)  ' Total Added
        wsInv.Cells(newRow, 7).Value = 0           ' Total Sold
        wsInv.Cells(newRow, 10).Value = Date       ' Last Updated
        
        ' Formulas for H and I are automatically handled by Excel Tables if formatted, but we ensure it:
        wsInv.Cells(newRow, 8).FormulaR1C1 = "=RC[-2]-RC[-1]"
        wsInv.Cells(newRow, 9).FormulaR1C1 = "=IF(RC[-1]<=0,""OUT OF STOCK"",IF(RC[-1]<=Settings!R11C2,""LOW STOCK"",""IN STOCK""))"
        
        ' Increment Settings Counter
        wsSet.Range("B7").Value = nextStkNum + 1
    Else
        ' EXISTING PRODUCT
        pID = wsInv.Cells(CLng(matchRow), 1).Value
        ' Update Total Added
        wsInv.Cells(CLng(matchRow), 6).Value = wsInv.Cells(CLng(matchRow), 6).Value + CDbl(pQty)
        ' Update Last Updated
        wsInv.Cells(CLng(matchRow), 10).Value = Date
    End If

    ' 4. Log to StockIn Ledger
    Dim ledgerRow As Long
    ledgerRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row + 1
    If ledgerRow < 13 Then ledgerRow = 13 ' Assuming table starts at row 12
    
    Dim stockRef As String
    stockRef = wsSet.Range("B10").Value & Format(wsSet.Range("B7").Value, "0000")
    
    wsStock.Cells(ledgerRow, 1).Value = stockRef
    wsStock.Cells(ledgerRow, 2).Value = Date
    wsStock.Cells(ledgerRow, 3).Value = pID
    wsStock.Cells(ledgerRow, 4).Value = pName
    wsStock.Cells(ledgerRow, 5).Value = pCat
    wsStock.Cells(ledgerRow, 6).Value = CDbl(pQty)
    wsStock.Cells(ledgerRow, 7).Value = CDbl(pCost)
    wsStock.Cells(ledgerRow, 8).FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-2]*RC[-1])"
    wsStock.Cells(ledgerRow, 9).Value = pUser

    ' 5. Refresh Dashboard
    Call RefreshDashboard

    ' 6. Clean Form
    wsStock.Range("D4:D8").ClearContents
    wsStock.Range("D9").Value = ""

    MsgBox "Stock updated successfully!" & vbNewLine & pName & " — Qty Added: " & pQty, vbInformation, "Stock Added"

    Exit Sub
ErrorHandler:
    MsgBox "An error occurred in AddStockItem: " & Err.Description, vbCritical
End Sub

Public Sub ProcessInvoice()
    On Error GoTo ErrorHandler
    
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets("Invoice")
    Dim wsStock As Worksheet: Set wsStock = ThisWorkbook.Sheets("Inventory")
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Sheets("Records")
    Dim wsSet As Worksheet: Set wsSet = ThisWorkbook.Sheets("Settings")

    ' --- Step 1: Validate ---
    If Trim(wsInv.Range("B12").Value) = "" Or Trim(wsInv.Range("B12").Value) = "Client / Company Name" Then
        MsgBox "Please enter a valid Client Name in B12.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    Dim r As Long
    Dim hasItems As Boolean: hasItems = False
    Dim oversellItem As String

    For r = 20 To 31
        If Trim(wsInv.Range("C" & r).Value) <> "" Then
            hasItems = True
            ' Check oversell
            Dim stockRemain As Variant
            stockRemain = wsInv.Range("H" & r).Value
            If IsNumeric(stockRemain) Then
                If stockRemain < 0 Then
                    oversellItem = wsInv.Range("C" & r).Value
                    MsgBox "Cannot process invoice! Insufficient stock for: " & oversellItem & vbNewLine & _
                           "Stock would fall below zero.", vbCritical, "Oversell Protection"
                    Exit Sub
                End If
            End If
        End If
    Next r

    If Not hasItems Then
        MsgBox "Please add at least one product to the invoice.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' --- Step 2: Capture Variables ---
    Dim invoiceID As String: invoiceID = wsInv.Range("G4").Value
    Dim clientName As String: clientName = wsInv.Range("B12").Value
    Dim clientComp As String: clientComp = wsInv.Range("B13").Value
    Dim grandTotal As Double: grandTotal = wsInv.Range("H36").Value
    Dim invDate As Date: invDate = wsInv.Range("G5").Value
    Dim itemCount As Long: itemCount = 0

    ' --- Step 3: Export PDF ---
    Dim pdfPath As String
    pdfPath = Environ("USERPROFILE") & "\Desktop\" & invoiceID & "_" & SafeFileName(clientName) & ".pdf"
    
    On Error Resume Next
    wsInv.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    If Err.Number <> 0 Then
        MsgBox "Failed to save PDF to Desktop. Close the PDF if it's already open.", vbCritical, "Export Error"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' --- Step 4: Deduct Stock ---
    For r = 20 To 31
        Dim pName As String
        pName = Trim(wsInv.Range("C" & r).Value)
        Dim pQty As Variant
        pQty = wsInv.Range("F" & r).Value
        
        If pName <> "" And IsNumeric(pQty) Then
            If pQty > 0 Then
                itemCount = itemCount + 1
                Dim matchRow As Variant
                matchRow = Application.Match(pName, wsStock.Columns("B"), 0)
                If Not IsError(matchRow) Then
                    ' Increment Total Sold (Col G)
                    wsStock.Cells(CLng(matchRow), 7).Value = wsStock.Cells(CLng(matchRow), 7).Value + CDbl(pQty)
                    wsStock.Cells(CLng(matchRow), 10).Value = Date
                End If
            End If
        End If
    Next r

    ' --- Step 5: Log to Records ---
    Dim nr As Long
    nr = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row + 1
    
    wsRec.Cells(nr, 1).Value = invoiceID
    wsRec.Cells(nr, 2).Value = invDate
    wsRec.Cells(nr, 3).Value = clientName
    wsRec.Cells(nr, 4).Value = clientComp
    wsRec.Cells(nr, 5).Value = grandTotal
    wsRec.Cells(nr, 6).Value = itemCount

    ' --- Step 6: Increment Counter ---
    wsSet.Range("B6").Value = wsSet.Range("B6").Value + 1

    ' --- Step 7: Clear Form ---
    Dim defaultLabels As Variant: defaultLabels = Array("Client Name", "Client Company", "Street Address", "Phone", "Email")
    Dim i As Integer
    For i = 0 To 4
        wsInv.Cells(12 + i, 2).Value = defaultLabels(i)
    Next i
    
    wsInv.Range("C20:F31").ClearContents
    wsInv.Range("H34").Value = 0
    wsInv.Range("G6").ClearContents

    ' --- Step 8: Refresh Dashboard ---
    Call RefreshDashboard

    ' --- Step 9: Confirm ---
    MsgBox "Invoice " & invoiceID & " processed successfully." & vbNewLine & _
           "PDF saved to Desktop." & vbNewLine & _
           "Stock levels updated.", vbInformation, "Success"

    Exit Sub
ErrorHandler:
    MsgBox "An error occurred in ProcessInvoice: " & Err.Description, vbCritical
End Sub

Public Sub RefreshDashboard()
    On Error Resume Next
    Dim wsDash As Worksheet: Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets("Inventory")
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Sheets("Records")
    
    ' Update KPI Formulas explicitly
    wsDash.Range("C6").Formula = "=COUNTA(Inventory!A2:A1000)"
    wsDash.Range("E6").Formula = "=SUMPRODUCT(Inventory!E2:E1000,Inventory!H2:H1000)"
    wsDash.Range("G6").Formula = "=COUNTIF(Inventory!I2:I1000,""LOW STOCK"")"
    wsDash.Range("I6").Formula = "=COUNTIF(Inventory!I2:I1000,""OUT OF STOCK"")"
    
    wsDash.Range("C11").Formula = "=COUNTA(Records!A2:A1000)"
    wsDash.Range("E11").Formula = "=SUM(Records!E2:E1000)"

    ' Refresh all Pivot Tables and Charts
    ThisWorkbook.RefreshAll
    DoEvents
    On Error GoTo 0
End Sub

Public Sub GoToDashboard()
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Activate
    On Error GoTo 0
End Sub

Private Function SafeFileName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "")
    Next c
    SafeFileName = Trim(s)
End Function
