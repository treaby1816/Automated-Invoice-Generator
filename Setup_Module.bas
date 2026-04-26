Option Explicit

Private Function Navy() As Long: Navy = RGB(31, 56, 100): End Function
Private Function Accent() As Long: Accent = RGB(68, 114, 196): End Function
Private Function LFill() As Long: LFill = RGB(217, 225, 242): End Function
Private Function EFill() As Long: EFill = RGB(221, 235, 247): End Function
Private Function Teal() As Long: Teal = RGB(47, 85, 58): End Function
Private Function N() As String: N = ChrW(8358): End Function

Public Sub SetupInvoiceSystem()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Visible = xlSheetVisible
        On Error GoTo 0
    Next ws
    
    Dim wI As Worksheet: Set wI = wb.Sheets(1)
    wI.Cells.Clear
    Dim shp As Shape
    For Each shp In wI.Shapes: shp.Delete: Next shp
    
    Do While wb.Sheets.Count > 1
        On Error Resume Next
        wb.Sheets(wb.Sheets.Count).Delete
        On Error GoTo 0
    Loop
    wI.Name = "Invoice"
    
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets.Add(Before:=wI): wsDash.Name = "Dashboard"
    Dim wsInv As Worksheet: Set wsInv = wb.Sheets.Add(After:=wI): wsInv.Name = "Inventory"
    Dim wsStk As Worksheet: Set wsStk = wb.Sheets.Add(After:=wsInv): wsStk.Name = "StockIn"
    Dim wsRec As Worksheet: Set wsRec = wb.Sheets.Add(After:=wsStk): wsRec.Name = "Records"
    Dim wsSet As Worksheet: Set wsSet = wb.Sheets.Add(After:=wsRec): wsSet.Name = "Settings"
    Dim wsLic As Worksheet: Set wsLic = wb.Sheets.Add(After:=wsSet): wsLic.Name = "LicenseData"
    
    Call BuildSettings(wsSet)
    Call BuildLicenseData(wsLic)
    Call BuildInventory(wsInv)
    Call BuildStockIn(wsStk)
    Call BuildRecords(wsRec)
    Call BuildInvoice(wI)
    Call BuildDashboard(wsDash)
    
    wsLic.Visible = xlSheetVeryHidden
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    On Error Resume Next
    Dim freezeSheets As Variant
    freezeSheets = Array("Inventory", "StockIn", "Records")
    Dim shName As Variant
    For Each shName In freezeSheets
        wb.Sheets(CStr(shName)).Activate
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
    Next shName
    On Error GoTo 0
    
    wsDash.Activate
    ActiveWindow.DisplayGridlines = False
    
    MsgBox "Invoice & Inventory Management System setup complete!", vbInformation, "Setup Complete"
End Sub

Private Sub BuildSettings(ws As Worksheet)
    With ws
        .Columns("A").ColumnWidth = 24: .Columns("B").ColumnWidth = 40
        .Columns("D").ColumnWidth = 20
        .Range("A1:B1").Merge: .Range("A1").Value = "COMPANY SETTINGS"
        .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 13
        .Range("A1").Font.Color = vbWhite: .Range("A1:B1").Interior.Color = Navy()
        .Range("A1").HorizontalAlignment = xlCenter
        
        Dim L As Variant: L = Array("Company Name", "Address Line 1", "Address Line 2", "Phone", "Email", "Website", "RC Number", "Bank Name", "Account Name", "Account Number", "Sort Code", "Default VAT (%)", "Invoice Prefix", "Next Invoice #", "Payment Terms (days)", "Stock-In Prefix", "Next Stock-In Reference", "Low Stock Threshold", "Currency Symbol")
        Dim V As Variant: V = Array("BeFed Catering & Events", "3, Ibiyemi Ajadi Close", "Abeokuta, Ogun, Nigeria", "+234 800 000 0000", "info@yourcompany.com", "www.yourcompany.com", "RC 123456", "First Bank of Nigeria", "Your Company Name", "1234567890", "011", 7.5, "INV-", 1, 30, "STK-", 1, 10, N())
        
        Dim i As Long
        For i = 0 To UBound(L)
            .Cells(i + 2, 1).Value = L(i): .Cells(i + 2, 1).Font.Bold = True
            .Cells(i + 2, 2).Value = V(i)
        Next i
        
        Dim lastRow As Long: lastRow = UBound(L) + 2
        .Range("A2:A" & lastRow).Interior.Color = vbWhite
        .Range("B2:B" & lastRow).Interior.Color = EFill()
        .Range("A2:B" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A2:B" & lastRow).RowHeight = 20
        
        .Range("B13").NumberFormat = "0.0"
        
        Dim btn As Object
        Set btn = .Buttons.Add(.Range("A" & lastRow + 2).Left, .Range("A" & lastRow + 2).Top, .Range("A" & lastRow + 2 & ":B" & lastRow + 2).Width, 25)
        btn.Caption = "Apply Settings"
        btn.OnAction = "ApplySettings"
        
        .Range("D1").Value = "CATEGORIES"
        .Range("D1").Font.Bold = True: .Range("D1").Font.Color = vbWhite: .Range("D1").Interior.Color = Navy()
        Dim cats As Variant: cats = Array("Food", "Drinks", "Rentals", "Service", "Decor", "Logistics", "Other")
        For i = 0 To UBound(cats)
            .Cells(i + 2, 4).Value = cats(i)
        Next i
        .Range("D2:D8").Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub BuildLicenseData(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Stored Key": .Cells(1, 2).Value = "WAITING_ACTIVATION"
        .Cells(2, 1).Value = "Company": .Cells(2, 2).Value = "Moresta Signature"
        .Cells(3, 1).Value = "Expiry Date": .Cells(3, 2).Value = "31-Dec-2099"
        .Cells(4, 1).Value = "Activation Date"
        .Cells(5, 1).Value = "Hardware Fingerprint"
        .Cells(6, 1).Value = "Grace Logins": .Cells(6, 2).Value = 3
    End With
End Sub

Private Sub BuildInventory(ws As Worksheet)
    With ws
        Dim h As Variant: h = Array("Product ID", "Product Name", "Category", "Description", "Unit Price (" & N() & ")", "Total Stock Added", "Total Stock Sold", "Current Stock", "Stock Status", "Last Updated")
        Dim w As Variant: w = Array(12, 25, 15, 20, 15, 15, 15, 15, 15, 12)
        Dim i As Long
        For i = 0 To 9
            .Cells(1, i + 1).Value = h(i): .Columns(i + 1).ColumnWidth = w(i)
        Next i
        
        With .Range("A1:J1")
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = Navy(): .Font.Size = 10
            .HorizontalAlignment = xlCenter
        End With
        
        Dim lo As ListObject
        Set lo = .ListObjects.Add(xlSrcRange, .Range("A1:J2"), , xlYes)
        lo.Name = "tblInventory"
        lo.TableStyle = "TableStyleMedium2"
        
        With .Range("C2:C1000").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Settings!$D$2:$D$8"
        End With
        
        .Range("E2:E1000").NumberFormat = "#,##0.00"
        Dim q As String: q = Chr(34)
        .Range("H2").Formula = "=IF([@[Product Name]]=" & q & q & "," & q & q & ",[@[Total Stock Added]]-[@[Total Stock Sold]])"
        .Range("I2").Formula = "=IF([@[Current Stock]]<=0," & q & "OUT OF STOCK" & q & ",IF([@[Current Stock]]<=Settings!$B$19," & q & "LOW STOCK" & q & "," & q & "IN STOCK" & q & "))"
        
        Dim fc As Object
        .Range("I2:I1000").FormatConditions.Delete
        Set fc = .Range("I2:I1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OUT OF STOCK""")
        fc.Interior.Color = RGB(220, 53, 69): fc.Font.Color = vbWhite: fc.Font.Bold = True
        Set fc = .Range("I2:I1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""LOW STOCK""")
        fc.Interior.Color = RGB(253, 126, 20): fc.Font.Color = vbBlack: fc.Font.Bold = True
        Set fc = .Range("I2:I1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""IN STOCK""")
        fc.Interior.Color = RGB(40, 167, 69): fc.Font.Color = vbWhite: fc.Font.Bold = True
    End With
End Sub

Private Sub BuildStockIn(ws As Worksheet)
    With ws
        .Columns("B").ColumnWidth = 4: .Columns("C").ColumnWidth = 20: .Columns("D").ColumnWidth = 30
        .Range("C2:D2").Merge: .Range("C2").Value = "ADD STOCK / NEW PRODUCT"
        .Range("C2").Font.Bold = True: .Range("C2").Font.Size = 14: .Range("C2").Font.Color = vbWhite
        .Range("C2").Interior.Color = Navy(): .Range("C2").HorizontalAlignment = xlCenter
        
        Dim fLabels As Variant: fLabels = Array("Product Name *", "Category *", "Description", "Quantity to Add *", "Unit Cost (" & N() & ") *", "Logged By")
        Dim i As Long
        For i = 0 To 5
            .Cells(i + 4, 3).Value = fLabels(i): .Cells(i + 4, 3).Font.Bold = True
            .Cells(i + 4, 4).Interior.Color = EFill()
            .Cells(i + 4, 4).Borders.LineStyle = xlContinuous
        Next i
        
        With .Range("D5").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Settings!$D$2:$D$8"
        End With
        
        Dim btn As Object
        Set btn = .Buttons.Add(.Range("C10").Left, .Range("C10").Top, .Range("C10:D10").Width, 30)
        btn.Caption = "ADD TO STOCK"
        btn.OnAction = "AddStockItem"
        btn.Font.Bold = True
        
        .Range("A13").Value = "Stock-In Ref"
        .Range("B13").Value = "Date"
        .Range("C13").Value = "Product ID"
        .Range("D13").Value = "Product Name"
        .Range("E13").Value = "Category"
        .Range("F13").Value = "Quantity Added"
        .Range("G13").Value = "Unit Cost (" & N() & ")"
        .Range("H13").Value = "Total Cost (" & N() & ")"
        .Range("I13").Value = "Logged By"
        
        With .Range("A13:I13")
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = Navy()
        End With
        
        Dim lo As ListObject
        Set lo = .ListObjects.Add(xlSrcRange, .Range("A13:I14"), , xlYes)
        lo.Name = "tblStockIn"
        lo.TableStyle = "TableStyleMedium2"
        
        .Columns("A:I").AutoFit
        .Columns("A").ColumnWidth = 14
    End With
End Sub

Private Sub BuildRecords(ws As Worksheet)
    With ws
        Dim h As Variant: h = Array("Invoice ID", "Date", "Client Name", "Client Company", "Grand Total (" & N() & ")", "Items Sold")
        Dim w As Variant: w = Array(15, 12, 25, 25, 20, 12)
        Dim i As Long
        For i = 0 To 5
            .Cells(1, i + 1).Value = h(i): .Columns(i + 1).ColumnWidth = w(i)
        Next i
        
        With .Range("A1:F1")
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = Navy()
        End With
        
        Dim lo As ListObject
        Set lo = .ListObjects.Add(xlSrcRange, .Range("A1:F2"), , xlYes)
        lo.Name = "tblRecords"
        lo.TableStyle = "TableStyleMedium2"
        
        .Columns("B").NumberFormat = "DD-MMM-YYYY"
        .Columns("E").NumberFormat = "#,##0.00"
    End With
End Sub

Private Sub BuildInvoice(ws As Worksheet)
    With ws
        .Cells.Clear
        Dim shpOld As Shape
        For Each shpOld In .Shapes: shpOld.Delete: Next shpOld

        .Columns("A").ColumnWidth = 3: .Columns("B").ColumnWidth = 27
        .Columns("C").ColumnWidth = 10: .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 16: .Columns("F").ColumnWidth = 18
        
        .Rows("1:2").RowHeight = 35
        .Range("A1:F2").Interior.Color = RGB(31, 73, 125)
        
        Dim shp As Shape
        Set shp = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("A1").Left + 10, .Range("A1").Top + 10, 60, 45)
        With shp
            .Name = "LogoPlaceholder"
            .TextFrame2.TextRange.Text = "ADD LOGO"
            .TextFrame2.TextRange.Font.Size = 8: .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(31, 73, 125)
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .Fill.ForeColor.RGB = vbWhite
            .Line.ForeColor.RGB = vbWhite: .Line.Weight = 2
            .OnAction = "InsertLogo"
        End With
        
        .Range("C1:E2").Merge: .Range("C1").Formula = "=Settings!B2"
        .Range("C1").Font.Color = vbWhite: .Range("C1").Font.Size = 22: .Range("C1").Font.Bold = True
        .Range("C1").HorizontalAlignment = xlCenter: .Range("C1").VerticalAlignment = xlCenter
        .Range("C1").ShrinkToFit = True
        
        .Range("F1:F2").Merge: .Range("F1").Value = "INVOICE"
        .Range("F1").Font.Color = vbWhite: .Range("F1").Font.Size = 24: .Range("F1").Font.Bold = True
        .Range("F1").HorizontalAlignment = xlRight: .Range("F1").VerticalAlignment = xlCenter
        
        .Range("A3:D3").Merge: .Range("A3").Formula = "=Settings!B3"
        .Range("A4:D4").Merge: .Range("A4").Formula = "=Settings!B4"
        .Range("A5:D5").Merge: .Range("A5").Formula = "=""Phone: "" & Settings!B5"
        .Range("A6:D6").Merge: .Range("A6").Formula = "=""RC Number: "" & Settings!B8"
        .Range("A3:A6").Font.Size = 10: .Range("A3:A6").Font.Color = RGB(80, 80, 80)
        
        Dim meta As Variant: meta = Array("Invoice #:", "Invoice Date:", "Due Date:", "Status:")
        Dim i As Integer
        For i = 0 To 3
            .Cells(3 + i, 5).Value = meta(i): .Cells(3 + i, 5).Font.Bold = True
            .Cells(3 + i, 5).HorizontalAlignment = xlRight: .Cells(3 + i, 5).Font.Size = 10
        Next i
        
        .Range("F3").Formula = "=Settings!B14&TEXT(Settings!B15,""0000"")"
        .Range("F4").Formula = "=TEXT(TODAY(),""DD/MM/YYYY"")"
        .Range("F5").Formula = "=TEXT(TODAY()+Settings!B16,""DD/MM/YYYY"")"
        .Range("F6").Value = "UNPAID": .Range("F6").Font.Color = vbRed: .Range("F6").Font.Bold = True
        .Range("F3:F6").HorizontalAlignment = xlRight: .Range("F3:F6").Font.Size = 10

        .Rows("7").RowHeight = 15

        .Rows("8").RowHeight = 18
        .Range("A8:D8").Merge: .Range("A8").Value = "BILL TO:": .Range("A8").Interior.Color = RGB(50, 120, 190)
        .Range("E8:F8").Merge: .Range("E8").Value = "SHIP TO:": .Range("E8").Interior.Color = RGB(50, 120, 190)
        .Range("A8, E8").Font.Color = vbWhite: .Range("A8, E8").Font.Bold = True: .Range("A8, E8").IndentLevel = 1
        
        Dim lb As Variant: lb = Array("Client / Company Name", "Street Address", "City, State", "Phone", "Email")
        For i = 0 To 4
            .Range(.Cells(9 + i, 1), .Cells(9 + i, 4)).Merge
            .Cells(9 + i, 1).Value = lb(i): .Cells(9 + i, 1).Font.Color = RGB(180, 180, 180): .Cells(9 + i, 1).Font.Italic = True
            .Range(.Cells(9 + i, 5), .Cells(9 + i, 6)).Merge
            .Cells(9 + i, 5).Value = lb(i): .Cells(9 + i, 5).Font.Color = RGB(180, 180, 180): .Cells(9 + i, 5).Font.Italic = True
        Next i

        .Rows("14:15").RowHeight = 10

        .Rows("16").RowHeight = 25
        .Range("A16:F16").Interior.Color = RGB(230, 230, 230)
        
        With .Range("A16:B16").Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 112, 192)
        End With
        With .Range("C16:D16").Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 176, 80)
        End With
        With .Range("E16").Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(255, 0, 0)
        End With
        With .Range("F16").Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(180, 180, 180)
        End With
        
        Dim btn As Object
        Dim tTop As Double: tTop = .Range("A16").Top + 4
        Dim tHeight As Double: tHeight = 19

        Set btn = ws.Buttons.Add(.Range("A16").Left, tTop, .Range("A16:B16").Width, tHeight)
        btn.Caption = "Print Invoice": btn.OnAction = "PrintInvoice"
        
        Set btn = ws.Buttons.Add(.Range("C16").Left, tTop, .Range("C16:D16").Width, tHeight)
        btn.Caption = "Save As PDF": btn.OnAction = "SaveAsPDF"
        
        Set btn = ws.Buttons.Add(.Range("E16").Left, tTop, .Range("E16").Width, tHeight)
        btn.Caption = "Clear Form": btn.OnAction = "ClearForm"
        
        Dim wSm As Double: wSm = (.Range("F16").Width / 2)
        
        Set btn = ws.Buttons.Add(.Range("F16").Left, tTop, wSm, tHeight)
        btn.Caption = "New Invoice": btn.OnAction = "NewInvoice"
        
        Set btn = ws.Buttons.Add(.Range("F16").Left + wSm, tTop, wSm, tHeight)
        btn.Caption = "Save Record": btn.OnAction = "SaveRecord"

        .Range("A17").Value = "#": .Range("B17:C17").Merge: .Range("B17").Value = "DESCRIPTION"
        .Range("D17").Value = "QTY": .Range("E17").Value = "UNIT PRICE (" & N() & ")": .Range("F17").Value = "AMOUNT (" & N() & ")"
        With .Range("A17:F17")
            .Interior.Color = RGB(31, 73, 125): .Font.Color = vbWhite: .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        Dim q2 As String: q2 = Chr(34)
        For i = 18 To 28
            .Cells(i, 1).Value = i - 17: .Cells(i, 1).HorizontalAlignment = xlCenter
            .Range("B" & i & ":C" & i).Merge
            
            With .Range("B" & i).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=INDIRECT(" & q2 & "tblInventory[Product Name]" & q2 & ")"
                .ErrorTitle = "Item Not in Inventory"
                .ErrorMessage = "This item is not in your Inventory. The price will not auto-fill. Please add it to the 'StockIn' sheet first."
                .ShowError = True
            End With
            
            .Cells(i, 4).Value = 0
            .Cells(i, 5).Formula = "=IF(B" & i & "=" & q2 & q2 & "," & q2 & q2 & ",IFERROR(INDEX(Inventory!E:E,MATCH(B" & i & ",Inventory!B:B,0)),0))"
            .Cells(i, 6).Formula = "=IF(OR(D" & i & "=0,E" & i & "=" & q2 & q2 & "),0,D" & i & "*E" & i & ")"
            If i Mod 2 = 0 Then .Range("A" & i & ":F" & i).Interior.Color = RGB(245, 248, 253)
        Next i

        Dim startTot As Integer: startTot = 29
        .Cells(startTot, 5).Value = "Subtotal (" & N() & "):": .Cells(startTot, 6).Formula = "=SUM(F18:F28)"
        .Cells(startTot + 1, 5).Value = "Discount (%):": .Cells(startTot + 1, 6).Value = 0
        .Cells(startTot + 2, 5).Value = "Discount (" & N() & "):": .Cells(startTot + 2, 6).Formula = "=F29*(F30/100)"
        .Cells(startTot + 3, 5).Value = "VAT / Tax (%):": .Cells(startTot + 3, 6).Formula = "=Settings!B13"
        .Cells(startTot + 4, 5).Value = "Tax Amount (" & N() & "):": .Cells(startTot + 4, 6).Formula = "=(F29-F31)*(F32/100)"
        
        .Range("A35:F35").Interior.Color = RGB(31, 73, 125)
        .Range("A35:E35").Merge: .Range("A35").Value = "TOTAL DUE (" & N() & "):"
        .Range("A35").Font.Color = vbWhite: .Range("A35").Font.Bold = True: .Range("A35").HorizontalAlignment = xlRight
        .Range("F35").Formula = "=F29-F31+F33"
        .Range("F35").Font.Color = vbWhite: .Range("F35").Font.Bold = True: .Range("F35").HorizontalAlignment = xlRight
        
        .Range("F29:F35").NumberFormat = "#,##0.00"

        .Range("A37:F37").Merge: .Range("A37").Value = "PAYMENT INSTRUCTIONS": .Range("A37").Interior.Color = RGB(50, 120, 190)
        .Range("A37").Font.Color = vbWhite: .Range("A37").Font.Bold = True: .Range("A37").IndentLevel = 1
        
        .Range("A38").Value = "Bank:": .Range("B38").Formula = "=Settings!B9"
        .Range("A39").Value = "A/C Name:": .Range("B39").Formula = "=Settings!B10"
        .Range("A40").Value = "A/C No:": .Range("B40").Formula = "=Settings!B11"
        .Range("A41").Value = "Sort:": .Range("B41").Formula = "=Settings!B12"
        .Range("A38:A41").Font.Bold = True
        
        .Range("A43:F43").Merge: .Range("A43").Value = "NOTES & TERMS": .Range("A43").Interior.Color = RGB(50, 120, 190)
        .Range("A43").Font.Color = vbWhite: .Range("A43").Font.Bold = True: .Range("A43").IndentLevel = 1
        
        .Range("A44").Formula = "=""1. Payment due within "" & Settings!B16 & "" days of invoice date."""
        .Range("A45").Value = "2. Late payments attract 2% monthly interest."
        .Range("A46").Value = "3. Disputes must be raised within 7 days of receipt."
        .Range("A44:A46").Font.Size = 9: .Range("A44:A46").Font.Color = RGB(60, 60, 60)
        
        .Range("A48:F48").Merge: .Range("A48").Value = "Thank you for your business!"
        .Range("A48").Font.Italic = True: .Range("A48").Font.Color = RGB(31, 73, 125): .Range("A48").Font.Size = 14: .Range("A48").HorizontalAlignment = xlCenter
        
        .Range("A49:F49").Merge: .Range("A49").Formula = "=""Generated by "" & Settings!B2 & "" Invoice System"""
        .Range("A49").Font.Size = 9: .Range("A49").Font.Color = RGB(180, 180, 180): .Range("A49").HorizontalAlignment = xlCenter
        
        .PageSetup.PrintArea = "$A$1:$F$50"
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.Zoom = False
    End With
End Sub

Private Sub BuildDashboard(ws As Worksheet)
    With ws
        .Tab.Color = RGB(21, 21, 21)
        .Columns("A").ColumnWidth = 2
        .Columns("B:L").ColumnWidth = 15
        .Range("B1:L3").Merge: .Range("B1").Value = "INVENTORY & SALES DASHBOARD"
        .Range("B1").Interior.Color = RGB(21, 21, 21): .Range("B1").Font.Color = vbWhite
        .Range("B1").Font.Size = 20: .Range("B1").Font.Bold = True
        .Range("B1").VerticalAlignment = xlCenter: .Range("B1").IndentLevel = 1
        
        Dim Lbl(0 To 5) As String
        Lbl(0) = "Total Products": Lbl(1) = "Total Stock Value": Lbl(2) = "Low Stock Items"
        Lbl(3) = "Out of Stock": Lbl(4) = "Total Invoices": Lbl(5) = "Total Revenue (" & N() & ")"
        
        Dim Frm(0 To 5) As String
        Frm(0) = "=COUNTA('Inventory'!A2:A1000)"
        Frm(1) = "=SUMPRODUCT('Inventory'!E2:E1000,'Inventory'!H2:H1000)"
        Frm(2) = "=COUNTIF('Inventory'!I2:I1000,""LOW STOCK"")"
        Frm(3) = "=COUNTIF('Inventory'!I2:I1000,""OUT OF STOCK"")"
        Frm(4) = "=COUNTA('Records'!A2:A1000)"
        Frm(5) = "=SUM('Records'!E2:E1000)"
        
        Dim i As Integer, rowBase As Integer, colBase As Integer
        For i = 0 To 5
            If i < 4 Then
                rowBase = 5: colBase = 2 + (i * 2)
            Else
                rowBase = 10: colBase = 2 + ((i - 4) * 2)
            End If
            
            .Range(.Cells(rowBase, colBase), .Cells(rowBase + 2, colBase + 1)).Merge
            With .Cells(rowBase, colBase)
                .Formula = Frm(i)
                .Font.Size = 22: .Font.Bold = True
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Interior.Color = RGB(245, 245, 245)
                .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(150, 150, 150)
                If i = 1 Or i = 5 Then .NumberFormat = N() & "#,##0"
            End With
            
            .Range(.Cells(rowBase + 3, colBase), .Cells(rowBase + 3, colBase + 1)).Merge
            With .Cells(rowBase + 3, colBase)
                .Value = Lbl(i)
                .Font.Size = 9: .Font.Color = RGB(100, 100, 100)
                .HorizontalAlignment = xlCenter
            End With
        Next i

        .Range("B32:L32").Merge: .Range("B32").Value = "ITEMS REQUIRING RESTOCKING"
        .Range("B32").Interior.Color = RGB(220, 53, 69): .Range("B32").Font.Color = vbWhite
        .Range("B32").Font.Bold = True: .Range("B32").VerticalAlignment = xlCenter
        
        .Range("B33").Value = "Run Dashboard to see low-stock items"
        .Range("B33").Font.Italic = True: .Range("B33").Font.Color = RGB(120, 120, 120)
    End With
End Sub
