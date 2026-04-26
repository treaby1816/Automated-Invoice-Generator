Option Explicit

Public Sub AddStockItem()
    On Error GoTo ErrorHandler

    Dim wsStock As Worksheet: Set wsStock = ThisWorkbook.Sheets("StockIn")
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets("Inventory")
    Dim wsSet As Worksheet: Set wsSet = ThisWorkbook.Sheets("Settings")

    Dim pName As String: pName = Trim(wsStock.Range("D4").Value)
    Dim pCat As String: pCat = Trim(wsStock.Range("D5").Value)
    Dim pDesc As String: pDesc = Trim(wsStock.Range("D6").Value)
    Dim pQty As Variant: pQty = wsStock.Range("D7").Value
    Dim pCost As Variant: pCost = wsStock.Range("D8").Value
    Dim pUser As String: pUser = Trim(wsStock.Range("D9").Value)

    If pName = "" Or pCat = "" Or Not IsNumeric(pQty) Or Not IsNumeric(pCost) Then
        MsgBox "Please fill all required fields correctly (*).", vbExclamation, "Validation Error"
        Exit Sub
    End If
    If pQty <= 0 Or pCost < 0 Then
        MsgBox "Quantity must be > 0 and Cost must be >= 0.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    Dim matchRow As Variant
    matchRow = Application.Match(pName, wsInv.Columns("B"), 0)
    
    Dim pID As String

    If IsError(matchRow) Then
        Dim nextStkNum As Long
        nextStkNum = wsSet.Range("B18").Value
        pID = wsSet.Range("B17").Value & Format(nextStkNum, "0000")
        
        Dim newRow As Long
        newRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row + 1
        
        wsInv.Cells(newRow, 1).Value = pID
        wsInv.Cells(newRow, 2).Value = pName
        wsInv.Cells(newRow, 3).Value = pCat
        wsInv.Cells(newRow, 4).Value = pDesc
        wsInv.Cells(newRow, 5).Value = CDbl(pCost)
        wsInv.Cells(newRow, 6).Value = CDbl(pQty)
        wsInv.Cells(newRow, 7).Value = 0
        wsInv.Cells(newRow, 10).Value = Date
        
        wsInv.Cells(newRow, 8).FormulaR1C1 = "=RC[-2]-RC[-1]"
        wsInv.Cells(newRow, 9).FormulaR1C1 = "=IF(RC[-1]<=0,""OUT OF STOCK"",IF(RC[-1]<=Settings!R11C2,""LOW STOCK"",""IN STOCK""))"
        
        wsSet.Range("B18").Value = nextStkNum + 1
    Else
        pID = wsInv.Cells(CLng(matchRow), 1).Value
        wsInv.Cells(CLng(matchRow), 6).Value = wsInv.Cells(CLng(matchRow), 6).Value + CDbl(pQty)
        wsInv.Cells(CLng(matchRow), 10).Value = Date
    End If

    Dim ledgerRow As Long
    ledgerRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row + 1
    If ledgerRow < 13 Then ledgerRow = 13
    
    Dim stockRef As String
    stockRef = wsSet.Range("B17").Value & Format(wsSet.Range("B18").Value - 1, "0000")
    
    wsStock.Cells(ledgerRow, 1).Value = stockRef
    wsStock.Cells(ledgerRow, 2).Value = Date
    wsStock.Cells(ledgerRow, 3).Value = pID
    wsStock.Cells(ledgerRow, 4).Value = pName
    wsStock.Cells(ledgerRow, 5).Value = pCat
    wsStock.Cells(ledgerRow, 6).Value = CDbl(pQty)
    wsStock.Cells(ledgerRow, 7).Value = CDbl(pCost)
    wsStock.Cells(ledgerRow, 8).FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-2]*RC[-1])"
    wsStock.Cells(ledgerRow, 9).Value = pUser

    Call RefreshDashboard

    wsStock.Range("D4:D8").ClearContents
    wsStock.Range("D9").Value = ""

    MsgBox "Stock updated successfully!" & vbNewLine & pName & " - Qty Added: " & pQty, vbInformation, "Stock Added"

    Exit Sub
ErrorHandler:
    MsgBox "An error occurred in AddStockItem: " & Err.Description, vbCritical
End Sub

Public Sub PrintInvoice()
    On Error Resume Next
    ThisWorkbook.Sheets("Invoice").PrintPreview
    On Error GoTo 0
End Sub

Public Sub SaveAsPDF()
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets("Invoice")
    Dim invoiceID As String: invoiceID = wsInv.Range("F2").Value
    Dim clientName As String: clientName = wsInv.Range("B10").Value
    If clientName = "" Or clientName = "Client / Company Name" Then clientName = "Unknown_Client"
    
    Dim pdfPath As String
    pdfPath = Environ("USERPROFILE") & "\Desktop\" & invoiceID & "_" & SafeFileName(clientName) & ".pdf"
    
    On Error Resume Next
    wsInv.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    If Err.Number = 0 Then
        MsgBox "PDF saved to Desktop!", vbInformation
    Else
        MsgBox "Failed to save PDF.", vbCritical
    End If
    On Error GoTo 0
End Sub

Public Sub ClearForm()
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets("Invoice")
    wsInv.Range("B18:B28").ClearContents
    wsInv.Range("D18:D28").ClearContents
    wsInv.Range("F30").Value = 0
    
    Dim defaultLabels As Variant: defaultLabels = Array("Client / Company Name", "Street Address", "City, State", "Phone", "Email")
    Dim i As Integer
    For i = 0 To 4
        wsInv.Cells(10 + i, 2).Value = defaultLabels(i)
    Next i
End Sub

Public Sub NewInvoice()
    Dim wsSet As Worksheet: Set wsSet = ThisWorkbook.Sheets("Settings")
    wsSet.Range("B15").Value = wsSet.Range("B15").Value + 1
    Call ClearForm
End Sub

Public Sub SaveRecord()
    On Error GoTo ErrorHandler
    Dim wsInv As Worksheet: Set wsInv = ThisWorkbook.Sheets("Invoice")
    Dim wsStock As Worksheet: Set wsStock = ThisWorkbook.Sheets("Inventory")
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Sheets("Records")
    Dim wsSet As Worksheet: Set wsSet = ThisWorkbook.Sheets("Settings")

    If Trim(wsInv.Range("B10").Value) = "" Or Trim(wsInv.Range("B10").Value) = "Client / Company Name" Then
        MsgBox "Please enter a valid Client Name in B10.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    Dim r As Long
    Dim hasItems As Boolean: hasItems = False
    For r = 18 To 28
        If Trim(wsInv.Range("B" & r).Value) <> "" Then
            hasItems = True
            If IsNumeric(wsInv.Range("G" & r).Value) Then
                If wsInv.Range("G" & r).Value < 0 Then
                    MsgBox "Insufficient stock for: " & wsInv.Range("B" & r).Value, vbCritical, "Oversell Protection"
                    Exit Sub
                End If
            End If
        End If
    Next r

    If Not hasItems Then
        MsgBox "Please add at least one product.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    Dim invoiceID As String: invoiceID = wsInv.Range("F2").Value
    
    ' Check if record already exists
    Dim matchRec As Variant
    matchRec = Application.Match(invoiceID, wsRec.Columns("A"), 0)
    If Not IsError(matchRec) Then
        MsgBox "This Invoice (" & invoiceID & ") has already been saved!", vbExclamation
        Exit Sub
    End If

    Dim itemCount As Long: itemCount = 0
    For r = 18 To 28
        Dim pName As String: pName = Trim(wsInv.Range("B" & r).Value)
        Dim pQty As Variant: pQty = wsInv.Range("D" & r).Value
        If pName <> "" And IsNumeric(pQty) Then
            If pQty > 0 Then
                itemCount = itemCount + 1
                Dim matchRow As Variant: matchRow = Application.Match(pName, wsStock.Columns("B"), 0)
                If Not IsError(matchRow) Then
                    wsStock.Cells(CLng(matchRow), 7).Value = wsStock.Cells(CLng(matchRow), 7).Value + CDbl(pQty)
                    wsStock.Cells(CLng(matchRow), 10).Value = Date
                End If
            End If
        End If
    Next r

    Dim nr As Long: nr = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row + 1
    wsRec.Cells(nr, 1).Value = invoiceID
    wsRec.Cells(nr, 2).Value = wsInv.Range("F3").Value
    wsRec.Cells(nr, 3).Value = wsInv.Range("B10").Value
    wsRec.Cells(nr, 4).Value = ""
    wsRec.Cells(nr, 5).Value = wsInv.Range("F35").Value
    wsRec.Cells(nr, 6).Value = itemCount

    Call RefreshDashboard
    MsgBox "Record saved and stock levels updated successfully!", vbInformation, "Success"
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub RefreshDashboard()
    On Error Resume Next
    Dim wsDash As Worksheet: Set wsDash = ThisWorkbook.Sheets("Dashboard")
    If wsDash Is Nothing Then Exit Sub
    
    wsDash.Range("B5").Formula = "=COUNTA('Inventory'!A2:A1000)"
    wsDash.Range("D5").Formula = "=SUMPRODUCT('Inventory'!E2:E1000,'Inventory'!H2:H1000)"
    wsDash.Range("F5").Formula = "=COUNTIF('Inventory'!I2:I1000,""LOW STOCK"")"
    wsDash.Range("H5").Formula = "=COUNTIF('Inventory'!I2:I1000,""OUT OF STOCK"")"
    wsDash.Range("B10").Formula = "=COUNTA('Records'!A2:A1000)"
    wsDash.Range("D10").Formula = "=SUM('Records'!E2:E1000)"

    Application.CalculateFull
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

Public Sub InsertLogo()
    Dim picPath As Variant
    picPath = Application.GetOpenFilename("Image Files (*.jpg; *.jpeg; *.png; *.bmp), *.jpg; *.jpeg; *.png; *.bmp", 1, "Select Your Company Logo")
    
    If picPath = False Then Exit Sub
    
    On Error Resume Next
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes("LogoPlaceholder")
    Dim sLeft As Double: sLeft = shp.Left
    Dim sTop As Double: sTop = shp.Top
    Dim sWidth As Double: sWidth = shp.Width
    Dim sHeight As Double: sHeight = shp.Height
    shp.Delete
    On Error GoTo 0
    
    Dim pic As Shape
    Set pic = ActiveSheet.Shapes.AddPicture(picPath, msoFalse, msoTrue, sLeft, sTop, -1, -1)
    pic.Name = "CompanyLogo"
    pic.LockAspectRatio = msoTrue
    
    If pic.Width / sWidth > pic.Height / sHeight Then
        pic.Width = sWidth
    Else
        pic.Height = sHeight
    End If
    
    pic.Left = sLeft + (sWidth - pic.Width) / 2
    pic.Top = sTop + (sHeight - pic.Height) / 2
    pic.OnAction = "InsertLogo"
End Sub

Public Sub ApplySettings()
    ThisWorkbook.RefreshAll
    Application.CalculateFullRebuild
    MsgBox "Settings updated successfully!", vbInformation, "Settings"
End Sub
