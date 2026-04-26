Attribute VB_Name = "Setup_Module"
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
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wI As Worksheet: Set wI = wb.Sheets(1)
    
    ' Clear Invoice first to prevent reference errors when deleting sheets
    wI.Cells.Clear: wI.Cells.ClearFormats
    Dim shp As Shape
    For Each shp In wI.Shapes: shp.Delete: Next shp
    
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop
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
    wsDash.Activate
    ActiveWindow.DisplayGridlines = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Invoice & Inventory Management System setup complete!" & vbCrLf & _
           "Please complete the Developer Deployment Checklist in Module2.", vbInformation, "Setup Complete"
End Sub

Private Sub BuildSettings(ws As Worksheet)
    With ws
        .Columns("A").ColumnWidth = 24: .Columns("B").ColumnWidth = 32
        .Columns("D").ColumnWidth = 20
        .Range("A1:B1").Merge: .Range("A1").Value = "COMPANY SETTINGS"
        .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 13
        .Range("A1").Font.Color = vbWhite: .Range("A1:B1").Interior.Color = Navy()
        .Range("A1").HorizontalAlignment = xlCenter
        
        Dim L As Variant: L = Array("Company Name", "Address Line 1", "Address Line 2", "Phone", "Email", "Next Invoice Number", "Next Stock-In Reference", "VAT Rate", "Invoice Prefix", "Stock-In Prefix", "Low Stock Threshold", "Currency Symbol")
        Dim V As Variant: V = Array("Moresta Signature", "Ondo City, Ondo State", "Nigeria", "+234 000 000 0000", "morestasignature@gmail.com", 1, 1, 0.075, "INV-", "STK-", 10, N())
        Dim i As Long
        For i = 0 To 11
            .Cells(i + 2, 1).Value = L(i): .Cells(i + 2, 1).Font.Bold = True
            .Cells(i + 2, 2).Value = V(i)
        Next i
        
        .Range("A2:A13").Interior.Color = RGB(245, 245, 245)
        .Range("B2:B13").Interior.Color = EFill()
        .Range("A2:B13").Borders.LineStyle = xlContinuous
        .Range("B7").NumberFormat = "0": .Range("B8").NumberFormat = "0.00"
        .Range("B12").NumberFormat = "0"
        
        ' Categories
        .Range("D1").Value = "CATEGORIES"
        .Range("D1").Font.Bold = True: .Range("D1").Font.Color = vbWhite: .Range("D1").Interior.Color = Navy()
        Dim cats As Variant: cats = Array("Electronics", "Clothing", "Food & Beverage", "Pharmaceuticals", "Stationery", "Furniture", "Cosmetics", "Automotive", "Agriculture", "Other")
        For i = 0 To UBound(cats)
            .Cells(i + 2, 4).Value = cats(i)
        Next i
        .Range("D2:D11").Borders.LineStyle = xlContinuous
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
        
        ' Create Table
        Dim lo As ListObject
        Set lo = .ListObjects.Add(xlSrcRange, .Range("A1:J2"), , xlYes)
        lo.Name = "tblInventory"
        lo.TableStyle = "TableStyleMedium2"
        
        ' Data Validation for Category
        With .Range("C2:C1000").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Settings!$D$2:$D$11"
        End With
        
        .Range("E2:E1000").NumberFormat = "#,##0.00"
        .Range("H2:H1000").Formula = "=IF([@[Product Name]]="""","""",[@[Total Stock Added]]-[@[Total Stock Sold]])"
        .Range("I2:I1000").Formula = "=IF([@[Current Stock]]<=0,""OUT OF STOCK"",IF([@[Current Stock]]<=Settings!$B$12,""LOW STOCK"",""IN STOCK""))"
        
        ' Conditional Formatting for Status
        Dim fc As FormatCondition
        .Range("I2:I1000").FormatConditions.Delete
        Set fc = .Range("I2:I1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OUT OF STOCK""")
        fc.Interior.Color = RGB(220, 53, 69): fc.Font.Color = vbWhite: fc.Font.Bold = True
        Set fc = .Range("I2:I1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""LOW STOCK""")
        fc.Interior.Color = RGB(253, 126, 20): fc.Font.Color = vbBlack: fc.Font.Bold = True
        Set fc = .Range("I2:I1000").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""IN STOCK""")
        fc.Interior.Color = RGB(40, 167, 69): fc.Font.Color = vbWhite: fc.Font.Bold = True
        
        ws.Activate
        ActiveWindow.FreezePanes = False
        .Rows("2:2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub BuildStockIn(ws As Worksheet)
    With ws
        ' --- IN-SHEET FORM FOR ADDING STOCK ---
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
        
        With .Range("D5").Validation ' Category Dropdown
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Settings!$D$2:$D$11"
        End With
        
        ' Add Stock Button
        Dim btn As Object
        Set btn = .Buttons.Add(.Range("C10").Left, .Range("C10").Top, .Range("C10:D10").Width, 30)
        btn.Caption = "ADD TO STOCK"
        btn.OnAction = "AddStockItem"
        btn.Font.Bold = True
        
        ' --- LEDGER TABLE ---
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
    ' This preserves the premium layout but integrates the VLOOKUPs and Stock Remaining column
    With ws
        .Columns("A").ColumnWidth = 4: .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 26: .Columns("D").ColumnWidth = 8
        .Columns("E").ColumnWidth = 12: .Columns("F").ColumnWidth = 14: .Columns("G").ColumnWidth = 14

        ' === TOP ACCENT LINE ===
        .Range("A1:G1").Interior.Color = Navy(): .Range("A1:G1").RowHeight = 5

        ' === LOGO PLACEHOLDER ===
        .Rows("2:3").RowHeight = 28
        Dim logoShape As Shape
        Set logoShape = .Shapes.AddShape(msoShapeRoundedRectangle, _
            .Range("A2").Left + 2, .Range("A2").Top + 2, 60, 50)
        With logoShape
            .Name = "LogoPlaceholder"
            .Fill.ForeColor.RGB = LFill()
            .Line.ForeColor.RGB = Accent()
            .Line.Weight = 1.5
            .TextFrame2.TextRange.Text = "LOGO"
            .TextFrame2.TextRange.Font.Size = 9
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Accent()
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = "InsertLogo"
        End With

        ' === HEADER ===
        .Range("B2:D3").Merge: .Range("B2").Formula = "=Settings!B2"
        .Range("B2").Font.Size = 20: .Range("B2").Font.Bold = True: .Range("B2").Font.Color = Navy()
        .Range("B2").VerticalAlignment = xlCenter: .Range("B2").HorizontalAlignment = xlCenter
        .Range("F2:G3").Merge: .Range("F2").Value = "INVOICE"
        .Range("F2").Font.Size = 26: .Range("F2").Font.Bold = True: .Range("F2").Font.Color = Accent()
        .Range("F2").HorizontalAlignment = xlRight: .Range("F2").VerticalAlignment = xlCenter

        .Range("B4").Formula = "=Settings!B3": .Range("B5").Formula = "=Settings!B4"
        .Range("B6").Formula = "=""Phone: ""&Settings!B5"
        .Range("B4:B6").Font.Size = 9: .Range("B4:B6").Font.Color = RGB(80, 80, 80)

        .Range("F4").Value = "Invoice #:": .Range("F5").Value = "Invoice Date:"
        .Range("F6").Value = "Due Date:": .Range("F7").Value = "Status:"
        .Range("F4:F7").Font.Bold = True: .Range("F4:F7").HorizontalAlignment = xlRight
        .Range("F4:F7").Font.Size = 10

        .Range("G4").Formula = "=Settings!B9&TEXT(Settings!B6,""0000"")"
        .Range("G4").Font.Bold = True: .Range("G4").Font.Color = Accent(): .Range("G4").Font.Size = 12
        .Range("G5").Formula = "=TEXT(TODAY(),""DD/MM/YYYY"")": .Range("G5").Font.Color = Accent()
        .Range("G6").NumberFormat = "DD/MM/YYYY": .Range("G6").Font.Color = Accent()
        .Range("G6").Interior.Color = RGB(255, 255, 153) ' Highlight due date
        .Range("G7").Value = "UNPAID": .Range("G7").Font.Color = RGB(220, 20, 20): .Range("G7").Font.Bold = True
        .Range("G4:G7").HorizontalAlignment = xlRight

        ' === BILL TO / SHIP TO ===
        .Range("A9:D9").Merge: .Range("A9").Value = "BILL TO:"
        .Range("E9:G9").Merge: .Range("E9").Value = "SHIP TO:"
        Dim lb As Variant: lb = Array("Client Name", "Client Company", "Street Address", "Phone", "Email")
        Dim j As Long
        For j = 0 To 4
            .Cells(11 + j, 2).Value = lb(j): .Cells(11 + j, 5).Value = lb(j)
            .Cells(11 + j, 2).Font.Color = RGB(180, 180, 180): .Cells(11 + j, 2).Font.Italic = True
            .Cells(11 + j, 5).Font.Color = RGB(180, 180, 180): .Cells(11 + j, 5).Font.Italic = True
        Next j

        ' === ITEMS TABLE (Row 19 header, 20-31 data) ===
        .Range("A19").Value = "#"
        .Range("B19:C19").Merge: .Range("B19").Value = "PRODUCT NAME"
        .Range("D19").Value = "QTY"
        .Range("E19").Value = "UNIT PRICE (" & N() & ")"
        .Range("F19").Value = "AMOUNT (" & N() & ")"
        .Range("G19").Value = "STOCK REMAINING"
        
        Dim r As Long
        For r = 20 To 31
            .Cells(r, 1).Value = r - 19: .Cells(r, 1).HorizontalAlignment = xlCenter
            .Range("B" & r & ":C" & r).Merge
            
            ' Validation for Product Name
            With .Range("B" & r).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=tblInventory[Product Name]"
            End With
            
            .Cells(r, 4).Value = 0
            ' Auto-fill Unit Price
            .Cells(r, 5).Formula = "=IF(B" & r & "="""","""",IFERROR(VLOOKUP(B" & r & ",tblInventory[[Product Name]:[Unit Price (" & N() & ")]],4,0),""""))"
            ' Auto Amount
            .Cells(r, 6).Formula = "=IF(OR(D" & r & "="""",E" & r & "="""",D" & r & "=0,E" & r & "=0),"""",D" & r & "*E" & r & ")"
            ' Live Stock Remaining
            .Cells(r, 7).Formula = "=IF(B" & r & "="""","""",IFERROR(VLOOKUP(B" & r & ",tblInventory[[Product Name]:[Current Stock]],7,0)-D" & r & ",""N/A""))"
            
            ' Conditional format Stock Remaining < 0
            Dim fc As FormatCondition
            .Cells(r, 7).FormatConditions.Delete
            Set fc = .Cells(r, 7).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            fc.Font.Color = vbRed: fc.Font.Bold = True
        Next r

        ' === TOTALS ===
        Dim c As String: c = N() & "#,##0.00"
        .Range("E33:F33").Merge: .Range("E33").Value = "Subtotal (" & N() & "):": .Range("E33").Font.Bold = True
        .Range("E33").HorizontalAlignment = xlRight
        .Range("G33").Formula = "=SUM(F20:F31)"
        
        .Range("E34:F34").Merge: .Range("E34").Value = "Discount (" & N() & "):": .Range("E34").Font.Bold = True
        .Range("E34").HorizontalAlignment = xlRight
        .Range("G34").Value = 0
        
        .Range("E35:F35").Merge: .Range("E35").Value = "VAT / Tax (" & N() & "):": .Range("E35").Font.Bold = True
        .Range("E35").HorizontalAlignment = xlRight
        .Range("G35").Formula = "=IF(G34="""",(G33)*Settings!B8,(G33-G34)*Settings!B8)"
        
        .Range("D36:F36").Merge: .Range("D36").Value = "GRAND TOTAL (" & N() & "):"
        .Range("D36").Font.Bold = True: .Range("D36").Font.Size = 13
        .Range("D36").HorizontalAlignment = xlRight
        .Range("G36").Formula = "=G33-G34+G35"
        .Range("G36").Font.Bold = True: .Range("G36").Font.Size = 13

        .Range("E20:G31").NumberFormat = c
        .Range("G33:G36").ShrinkToFit = True
        .Range("G33").NumberFormat = c: .Range("G34").NumberFormat = c
        .Range("G35").NumberFormat = c: .Range("G36").NumberFormat = c

        ' === BUTTONS ===
        Dim btn As Object
        Set btn = ws.Buttons.Add(.Range("A33").Left, .Range("A33").Top, .Range("A33:B34").Width, 30)
        btn.Caption = "PROCESS INVOICE": btn.OnAction = "ProcessInvoice": btn.Font.Bold = True
        
        Set btn = ws.Buttons.Add(.Range("C33").Left, .Range("C33").Top, .Range("C33").Width, 30)
        btn.Caption = "VIEW DASHBOARD": btn.OnAction = "GoToDashboard": btn.Font.Bold = True

        ' === FOOTER ===
        .Range("A50:G50").Merge: .Range("A50").Value = "Thank you for your business!"
        .Range("A50").Font.Italic = True: .Range("A50").Font.Color = Accent()
        .Range("A50").Font.Size = 12: .Range("A50").HorizontalAlignment = xlCenter
        .Range("A51:G51").Merge: .Range("A51").Formula = "=""Generated by ""&Settings!B2&"" Invoice System"""
        .Range("A51").Font.Size = 8: .Range("A51").Font.Color = RGB(170, 170, 170)
        .Range("A51").Font.Italic = True: .Range("A51").HorizontalAlignment = xlCenter

        ' FORMATTING
        .Range("A2:G7").Interior.Color = RGB(245, 245, 248)
        .Range("A8:G8").Interior.Color = Navy(): .Range("A8:G8").RowHeight = 3
        Dim rng As Variant
        For Each rng In Array(.Range("A9:D9"), .Range("E9:G9"))
            rng.Interior.Color = Navy(): rng.Font.Color = vbWhite
            rng.Font.Bold = True: rng.Font.Size = 11
        Next rng
        .Range("A10:G16").Borders.LineStyle = xlContinuous
        .Range("A10:G16").Borders.Color = RGB(215, 215, 215)
        .Range("A19:G19").Interior.Color = RGB(31, 31, 31): .Range("A19:G19").Font.Color = vbWhite
        .Range("A19:G19").Font.Bold = True: .Range("A19:G19").Font.Size = 10
        For r = 20 To 31
            If r Mod 2 = 0 Then .Range("A" & r & ":G" & r).Interior.Color = RGB(255, 255, 255) Else .Range("A" & r & ":G" & r).Interior.Color = LFill()
            .Range("A" & r & ":G" & r).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("A" & r & ":G" & r).Borders(xlEdgeBottom).Color = RGB(215, 215, 215)
        Next r
        .Range("E33:G36").Interior.Color = LFill()
        .Range("D36:G36").Interior.Color = RGB(21, 21, 21): .Range("D36:G36").Font.Color = vbWhite
        .PageSetup.PrintArea = "$A$1:$G$52": .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1: .PageSetup.FitToPagesTall = 1
    End With
End Sub

Private Sub BuildDashboard(ws As Worksheet)
    With ws
        .Tab.Color = RGB(21, 21, 21)
        .Columns("A").ColumnWidth = 2
        .Range("B1:L3").Merge: .Range("B1").Value = "📊 INVENTORY & SALES DASHBOARD"
        .Range("B1").Interior.Color = RGB(21, 21, 21): .Range("B1").Font.Color = vbWhite
        .Range("B1").Font.Size = 20: .Range("B1").Font.Bold = True
        .Range("B1").VerticalAlignment = xlCenter: .Range("B1").IndentLevel = 1
        
        ' KPI Cards
        Dim Lbl As Variant: Lbl = Array("Total Products", "Total Stock Value", "Low Stock Items", "Out of Stock", "Total Invoices", "Total Revenue (" & N() & ")")
        Dim Frm As Variant: Frm = Array("=COUNTA(Inventory!A2:A1000)", "=SUMPRODUCT(Inventory!E2:E1000,Inventory!H2:H1000)", "=COUNTIF(Inventory!I2:I1000,""LOW STOCK"")", "=COUNTIF(Inventory!I2:I1000,""OUT OF STOCK"")", "=COUNTA(Records!A2:A1000)", "=SUM(Records!E2:E1000)")
        
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
                If i = 1 Or i = 5 Then .NumberFormat = N() & "#,##0" ' Currency
            End With
            
            .Range(.Cells(rowBase + 3, colBase), .Cells(rowBase + 3, colBase + 1)).Merge
            With .Cells(rowBase + 3, colBase)
                .Value = Lbl(i)
                .Font.Size = 9: .Font.Color = RGB(100, 100, 100)
                .HorizontalAlignment = xlCenter
            End With
        Next i

        ' Low Stock Table
        .Range("B32:L32").Merge: .Range("B32").Value = "⚠️ ITEMS REQUIRING RESTOCKING"
        .Range("B32").Interior.Color = RGB(220, 53, 69): .Range("B32").Font.Color = vbWhite
        .Range("B32").Font.Bold = True: .Range("B32").VerticalAlignment = xlCenter
        
        .Range("B33").Formula2 = "=FILTER(tblInventory[[Product ID]:[Stock Status]],(tblInventory[Stock Status]=""LOW STOCK"")+(tblInventory[Stock Status]=""OUT OF STOCK""),""All stock levels healthy!"")"
        
        ' Note: Generating complex PivotCharts via VBA in an empty workbook often causes runtime errors
        ' because the PivotCaches cannot be generated without data.
        ' Instead, we have built the layout, KPIs, and Low Stock filter natively.
        ' The user can insert the 4 native charts manually once data is populated.
    End With
End Sub
