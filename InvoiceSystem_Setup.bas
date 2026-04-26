Attribute VB_Name = "modSetup"
Option Explicit

Private Function Navy() As Long: Navy = RGB(31, 56, 100): End Function
Private Function Accent() As Long: Accent = RGB(68, 114, 196): End Function
Private Function LFill() As Long: LFill = RGB(217, 225, 242): End Function
Private Function EFill() As Long: EFill = RGB(221, 235, 247): End Function
Private Function Teal() As Long: Teal = RGB(47, 85, 58): End Function
Private Function N() As String: N = ChrW(8358): End Function

Public Sub SetupInvoiceSystem()
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    Dim wb As Workbook: Set wb = ThisWorkbook
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop
    wb.Sheets(1).Name = "Invoice"
    Dim wS As Worksheet: Set wS = wb.Sheets.Add(After:=wb.Sheets("Invoice")): wS.Name = "Settings"
    Dim wR As Worksheet: Set wR = wb.Sheets.Add(After:=wS): wR.Name = "Records"
    Call BuildSettings(wS): Call BuildRecords(wR)
    Dim wI As Worksheet: Set wI = wb.Sheets("Invoice")
    Call BuildInvoice(wI): Call FormatInvoice(wI): Call MakeButtons(wI)
    wI.Activate: wI.Range("B10").Select: ActiveWindow.DisplayGridlines = False
    Application.DisplayAlerts = True: Application.ScreenUpdating = True
    MsgBox "Invoice System setup complete!" & vbCrLf & vbCrLf & "Click on the 'LOGO' placeholder to insert your brand logo.", vbInformation, "Setup Complete"
End Sub

Private Sub BuildSettings(ws As Worksheet)
    With ws
        .Columns("A").ColumnWidth = 24: .Columns("B").ColumnWidth = 32
        .Range("A1:B1").Merge: .Range("A1").Value = "COMPANY SETTINGS"
        .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 13
        .Range("A1").Font.Color = vbWhite: .Range("A1:B1").Interior.Color = Navy()
        .Range("A1").HorizontalAlignment = xlCenter
        Dim L As Variant: L = Array("Company Name", "Address Line 1", "Address Line 2", "Phone", "Email", "Website", "RC Number", "Bank Name", "Account Name", "Account Number", "Sort Code", "Default VAT (%)", "Invoice Prefix", "Next Invoice #", "Payment Terms (days)")
        Dim V As Variant: V = Array("Moresta Signature", "Ondo City, Ondo State", "Nigeria", "+234 000 000 0000", "morestasignature@gmail.com", "www.morestasignature.com", "RC 000000", "First Bank of Nigeria", "Moresta Signature", "0000000000", "000", 7.5, "INV-", 1, 30)
        Dim i As Long
        For i = 0 To 14
            .Cells(i + 2, 1).Value = L(i): .Cells(i + 2, 1).Font.Bold = True
            .Cells(i + 2, 2).Value = V(i)
        Next i
        .Range("A2:A16").Interior.Color = RGB(245, 245, 245)
        .Range("B2:B16").Interior.Color = EFill()
        .Range("A2:B16").Borders.LineStyle = xlContinuous
        .Range("A2:B16").Borders.Color = RGB(200, 200, 200)
        .Range("B15").NumberFormat = "0": .Range("B13").NumberFormat = "0.0"
        On Error Resume Next: ThisWorkbook.Names("NextInvNum").Delete: On Error GoTo 0
        ThisWorkbook.Names.Add Name:="NextInvNum", RefersTo:="=Settings!$B$15"
        Dim b As Object: Set b = .Buttons.Add(.Range("A18").Left, .Range("A18").Top, .Range("A18:B19").Width, .Range("A18:B19").Height)
        b.Caption = "Apply Settings": b.OnAction = "ApplySettings": b.Font.Bold = True
    End With
End Sub

Private Sub BuildRecords(ws As Worksheet)
    With ws
        Dim h As Variant: h = Array("Invoice #", "Date", "Client Name", "Client Email", "Subtotal (" & N() & ")", "Discount (%)", "Tax (%)", "Total (" & N() & ")", "Status")
        Dim w As Variant: w = Array(14, 13, 26, 24, 16, 12, 10, 16, 11)
        Dim i As Long
        For i = 0 To 8
            .Cells(1, i + 1).Value = h(i): .Columns(i + 1).ColumnWidth = w(i)
        Next i
        With .Range("A1:I1")
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = Navy(): .Font.Size = 10
            .WrapText = True: .RowHeight = 30: .HorizontalAlignment = xlCenter
        End With
        .Columns("B").NumberFormat = "DD/MM/YYYY"
        .Columns("E").NumberFormat = "#,##0.00": .Columns("H").NumberFormat = "#,##0.00"
    End With
End Sub

Private Sub BuildInvoice(ws As Worksheet)
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
        .Range("B2").VerticalAlignment = xlCenter: .Range("B2").IndentLevel = 1
        .Range("F2:G3").Merge: .Range("F2").Value = "INVOICE"
        .Range("F2").Font.Size = 26: .Range("F2").Font.Bold = True: .Range("F2").Font.Color = Accent()
        .Range("F2").HorizontalAlignment = xlRight: .Range("F2").VerticalAlignment = xlCenter

        .Range("B4").Formula = "=Settings!B3": .Range("B5").Formula = "=Settings!B4"
        .Range("B6").Formula = "=""Phone: ""&Settings!B5"
        .Range("B7").Formula = "=""RC Number: ""&Settings!B8"
        .Range("B4:B7").Font.Size = 9: .Range("B4:B7").Font.Color = RGB(80, 80, 80)

        .Range("F4").Value = "Invoice #:": .Range("F5").Value = "Invoice Date:"
        .Range("F6").Value = "Due Date:": .Range("F7").Value = "Status:"
        .Range("F4:F7").Font.Bold = True: .Range("F4:F7").HorizontalAlignment = xlRight
        .Range("F4:F7").Font.Size = 10

        .Range("G4").Formula = "=Settings!B14&TEXT(Settings!B15,""0000"")"
        .Range("G4").Font.Bold = True: .Range("G4").Font.Color = Accent(): .Range("G4").Font.Size = 12
        .Range("G5").Formula = "=TEXT(TODAY(),""DD/MM/YYYY"")": .Range("G5").Font.Color = Accent()
        .Range("G6").NumberFormat = "DD/MM/YYYY": .Range("G6").Font.Color = Accent()
        .Range("G7").Value = "UNPAID": .Range("G7").Font.Color = RGB(220, 20, 20): .Range("G7").Font.Bold = True
        .Range("G4:G7").HorizontalAlignment = xlRight

        ' === BILL TO / SHIP TO ===
        .Range("A9:D9").Merge: .Range("A9").Value = "BILL TO:"
        .Range("E9:G9").Merge: .Range("E9").Value = "SHIP TO:"
        Dim lb As Variant: lb = Array("Client / Company Name", "Street Address", "City, State", "Phone", "Email")
        Dim j As Long
        For j = 0 To 4
            .Cells(10 + j, 2).Value = lb(j): .Cells(10 + j, 5).Value = lb(j)
            .Cells(10 + j, 2).Font.Color = RGB(180, 180, 180): .Cells(10 + j, 2).Font.Italic = True
            .Cells(10 + j, 5).Font.Color = RGB(180, 180, 180): .Cells(10 + j, 5).Font.Italic = True
        Next j

        ' === ITEMS TABLE (Row 17 header, 18-29 data) ===
        .Range("A17").Value = "#": .Range("B17:C17").Merge: .Range("B17").Value = "DESCRIPTION"
        .Range("D17").Value = "QTY": .Range("E17").Value = "UNIT PRICE (" & N() & ")"
        .Range("F17:G17").Merge: .Range("F17").Value = "AMOUNT (" & N() & ")"
        Dim r As Long
        For r = 18 To 29
            .Cells(r, 1).Value = r - 17: .Cells(r, 1).HorizontalAlignment = xlCenter
            .Range("B" & r & ":C" & r).Merge
            .Range("F" & r & ":G" & r).Merge
            .Cells(r, 4).Value = 0: .Cells(r, 5).Value = 0
            .Cells(r, 6).Formula = "=IF(D" & r & "*E" & r & "=0,""-"",D" & r & "*E" & r & ")"
        Next r

        ' === TOTALS ===
        Dim c As String: c = N() & "#,##0.00"
        .Range("E31:F31").Merge: .Range("E31").Value = "Subtotal (" & N() & "):": .Range("E31").Font.Bold = True
        .Range("E31").HorizontalAlignment = xlRight
        .Range("G31").Formula = "=SUM(F18:F29)"
        .Range("E33:F33").Merge: .Range("E33").Value = "Discount (%):": .Range("E33").Font.Bold = True
        .Range("E33").HorizontalAlignment = xlRight: .Range("G33").Value = 0
        .Range("E34:F34").Merge: .Range("E34").Value = "Discount (" & N() & "):": .Range("E34").Font.Bold = True
        .Range("E34").HorizontalAlignment = xlRight
        .Range("G34").Formula = "=IF(G33=0,""-"",G31*G33/100)"
        .Range("E35:F35").Merge: .Range("E35").Value = "VAT / Tax (%):": .Range("E35").Font.Bold = True
        .Range("E35").HorizontalAlignment = xlRight
        .Range("G35").Formula = "=Settings!B13"
        .Range("E36:F36").Merge: .Range("E36").Value = "Tax Amount (" & N() & "):": .Range("E36").Font.Bold = True
        .Range("E36").HorizontalAlignment = xlRight
        .Range("G36").Formula = "=IF(G34=""-"",G31*G35/100,(G31-G34)*G35/100)"
        .Range("D37:F37").Merge: .Range("D37").Value = "TOTAL DUE (" & N() & "):"
        .Range("D37").Font.Bold = True: .Range("D37").Font.Size = 13
        .Range("D37").HorizontalAlignment = xlRight
        .Range("G37").Formula = "=IF(G34=""-"",G31+G36,G31-G34+G36)"
        .Range("G37").Font.Bold = True: .Range("G37").Font.Size = 13

        .Range("E18:E29").NumberFormat = c: .Range("F18:F29").NumberFormat = c
        .Range("G31").NumberFormat = c: .Range("G34").NumberFormat = c
        .Range("G36").NumberFormat = c: .Range("G37").NumberFormat = c
        .Range("G33").NumberFormat = "0.0": .Range("G35").NumberFormat = "0.0"

        ' === PAYMENT INSTRUCTIONS ===
        .Range("A39:G39").Merge: .Range("A39").Value = "PAYMENT INSTRUCTIONS"
        .Range("B40").Value = "Bank:": .Range("C40").Formula = "=Settings!B9"
        .Range("B41").Value = "A/C Name:": .Range("C41").Formula = "=Settings!B10"
        .Range("B42").Value = "A/C No:": .Range("C42").Formula = "=Settings!B11"
        .Range("B43").Value = "Sort:": .Range("C43").Formula = "=Settings!B12"
        .Range("B40:B43").Font.Bold = True: .Range("B40:B43").Font.Size = 10

        ' === NOTES & TERMS ===
        .Range("A45:G45").Merge: .Range("A45").Value = "NOTES & TERMS"
        .Range("B46").Value = "1. Payment due within 30 days of invoice date."
        .Range("B47").Value = "2. Late payments attract 2% monthly interest."
        .Range("B48").Value = "3. Disputes must be raised within 7 days of receipt."
        .Range("B46:B48").Font.Size = 9

        ' === FOOTER ===
        .Range("A50:G50").Merge: .Range("A50").Value = "Thank you for your business!"
        .Range("A50").Font.Italic = True: .Range("A50").Font.Color = Accent()
        .Range("A50").Font.Size = 12: .Range("A50").HorizontalAlignment = xlCenter

        .Range("A51:G51").Merge
        .Range("A51").Value = "Generated by Moresta Signature Invoice System"
        .Range("A51").Font.Size = 8: .Range("A51").Font.Color = RGB(170, 170, 170)
        .Range("A51").Font.Italic = True: .Range("A51").HorizontalAlignment = xlCenter

        .PageSetup.PrintArea = "$A$1:$G$52"
        .PageSetup.Zoom = False
        .PageSetup.Orientation = xlPortrait
        .PageSetup.FitToPagesWide = 1: .PageSetup.FitToPagesTall = 1
        .PageSetup.TopMargin = Application.InchesToPoints(0.3)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.3)
        .PageSetup.LeftMargin = Application.InchesToPoints(0.3)
        .PageSetup.RightMargin = Application.InchesToPoints(0.3)
    End With
End Sub

Private Sub FormatInvoice(ws As Worksheet)
    With ws
        .Range("A2:G7").Interior.Color = RGB(245, 245, 248)
        .Range("A8:G8").Interior.Color = Navy(): .Range("A8:G8").RowHeight = 3

        Dim rng As Variant
        For Each rng In Array(.Range("A9:D9"), .Range("E9:G9"))
            rng.Interior.Color = Navy(): rng.Font.Color = vbWhite
            rng.Font.Bold = True: rng.Font.Size = 11
        Next rng
        .Range("A10:D14").Interior.Color = RGB(250, 250, 252)
        .Range("E10:G14").Interior.Color = RGB(250, 250, 252)
        .Range("A10:G14").Borders.LineStyle = xlContinuous
        .Range("A10:G14").Borders.Color = RGB(215, 215, 215)

        With .Range("A17:G17")
            .Interior.Color = Navy(): .Font.Color = vbWhite
            .Font.Bold = True: .Font.Size = 10: .HorizontalAlignment = xlCenter
        End With
        Dim r As Long
        For r = 18 To 29
            If r Mod 2 = 0 Then
                .Range("A" & r & ":G" & r).Interior.Color = RGB(255, 255, 255)
            Else
                .Range("A" & r & ":G" & r).Interior.Color = LFill()
            End If
            .Range("A" & r & ":G" & r).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("A" & r & ":G" & r).Borders(xlEdgeBottom).Color = RGB(215, 215, 215)
            .Range("D" & r & ":E" & r).Interior.Color = EFill()
        Next r
        .Range("A18:A29").Font.Bold = True: .Range("A18:A29").Font.Color = Navy()

        .Range("E31:G31").Interior.Color = RGB(245, 245, 248)
        .Range("E31:G31").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("E33:G36").Interior.Color = LFill()
        .Range("E33:G36").Borders.LineStyle = xlContinuous
        .Range("E33:G36").Borders.Color = RGB(200, 200, 200)
        .Range("G33").Interior.Color = EFill()

        .Range("D37:G37").Interior.Color = Navy()
        .Range("D37:G37").Font.Color = vbWhite

        With .Range("A39:G39")
            .Interior.Color = Accent(): .Font.Color = vbWhite
            .Font.Bold = True: .Font.Size = 10
        End With
        .Range("A40:G43").Interior.Color = RGB(248, 248, 250)

        With .Range("A45:G45")
            .Interior.Color = Accent(): .Font.Color = vbWhite
            .Font.Bold = True: .Font.Size = 10
        End With
        .Range("A46:G48").Interior.Color = RGB(248, 248, 250)

        .Range("A52:G52").Interior.Color = Navy(): .Range("A52:G52").RowHeight = 4
    End With
End Sub

Private Sub MakeButtons(ws As Worksheet)
    Dim b As Object
    Set b = ws.Buttons.Add(ws.Range("A16").Left, ws.Range("A16").Top, ws.Range("A16:B16").Width, ws.Range("A16:B16").Height)
    b.Caption = "Print Invoice": b.OnAction = "PrintInvoice": b.Font.Bold = True: b.Font.Size = 9
    Set b = ws.Buttons.Add(ws.Range("C16").Left, ws.Range("C16").Top, ws.Range("C16").Width, ws.Range("C16").Height)
    b.Caption = "Save As PDF": b.OnAction = "SaveAsPDF": b.Font.Bold = True: b.Font.Size = 9
    Set b = ws.Buttons.Add(ws.Range("D16").Left, ws.Range("D16").Top, ws.Range("D16").Width, ws.Range("D16").Height)
    b.Caption = "Clear Form": b.OnAction = "ClearForm": b.Font.Bold = True: b.Font.Size = 9
    Set b = ws.Buttons.Add(ws.Range("E16").Left, ws.Range("E16").Top, ws.Range("E16").Width, ws.Range("E16").Height)
    b.Caption = "New Invoice": b.OnAction = "NewInvoice": b.Font.Bold = True: b.Font.Size = 9
    Set b = ws.Buttons.Add(ws.Range("F16").Left, ws.Range("F16").Top, ws.Range("F16:G16").Width, ws.Range("F16:G16").Height)
    b.Caption = "Save Record": b.OnAction = "SaveRecord": b.Font.Bold = True: b.Font.Size = 9
End Sub

'===============================================================================
' BUTTON MACROS
'===============================================================================
Public Sub PrintInvoice()
    On Error GoTo E
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Invoice")
    If Not Valid(ws) Then Exit Sub
    ws.PrintPreview: Exit Sub
E: MsgBox "Print error: " & Err.Description, vbCritical
End Sub

Public Sub SaveAsPDF()
    On Error GoTo E
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Invoice")
    If Not Valid(ws) Then Exit Sub
    On Error Resume Next: ThisWorkbook.Save: On Error GoTo E
    Dim fn As String: fn = ws.Range("G4").Value & "_" & SF(ws.Range("B10").Value)
    Dim sp As Variant
    sp = Application.GetSaveAsFilename(InitialFileName:=fn, FileFilter:="PDF Files (*.pdf), *.pdf", Title:="Save Invoice as PDF")
    If sp = False Then MsgBox "Cancelled.", vbInformation: Exit Sub
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(sp), Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF saved!", vbInformation: Exit Sub
E: MsgBox "PDF error: " & Err.Description, vbCritical
End Sub

Public Sub ClearForm()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Invoice")
    Dim lb As Variant: lb = Array("Client / Company Name", "Street Address", "City, State", "Phone", "Email")
    Dim j As Long
    For j = 0 To 4
        ws.Cells(10 + j, 2).Value = lb(j): ws.Cells(10 + j, 2).Font.Color = RGB(180, 180, 180): ws.Cells(10 + j, 2).Font.Italic = True
        ws.Cells(10 + j, 5).Value = lb(j): ws.Cells(10 + j, 5).Font.Color = RGB(180, 180, 180): ws.Cells(10 + j, 5).Font.Italic = True
    Next j
    Dim r As Long
    For r = 18 To 29
        ws.Range("B" & r & ":C" & r).ClearContents
        ws.Cells(r, 4).Value = 0: ws.Cells(r, 5).Value = 0
    Next r
    ws.Range("G33").Value = 0: ws.Range("G6").ClearContents
    ws.Range("G7").Value = "UNPAID": ws.Range("G7").Font.Color = RGB(220, 20, 20)
    ws.Range("B10").Select
End Sub

Public Sub NewInvoice()
    Call ClearForm
    Dim wS As Worksheet: Set wS = ThisWorkbook.Sheets("Settings")
    wS.Range("B15").Value = wS.Range("B15").Value + 1
    MsgBox "New invoice " & ThisWorkbook.Sheets("Invoice").Range("G4").Value & " ready.", vbInformation
End Sub

Public Sub SaveRecord()
    On Error GoTo E
    Dim wI As Worksheet: Set wI = ThisWorkbook.Sheets("Invoice")
    Dim wR As Worksheet: Set wR = ThisWorkbook.Sheets("Records")
    If Not Valid(wI) Then Exit Sub
    Dim nr As Long: nr = wR.Cells(wR.Rows.Count, "A").End(xlUp).Row + 1
    wR.Cells(nr, 1).Value = wI.Range("G4").Value
    wR.Cells(nr, 2).Value = Date: wR.Cells(nr, 2).NumberFormat = "DD/MM/YYYY"
    wR.Cells(nr, 3).Value = wI.Range("B10").Value
    wR.Cells(nr, 4).Value = wI.Range("B14").Value
    wR.Cells(nr, 5).Value = wI.Range("G31").Value: wR.Cells(nr, 5).NumberFormat = "#,##0.00"
    wR.Cells(nr, 6).Value = wI.Range("G33").Value
    wR.Cells(nr, 7).Value = wI.Range("G35").Value
    wR.Cells(nr, 8).Value = wI.Range("G37").Value: wR.Cells(nr, 8).NumberFormat = "#,##0.00"
    wR.Cells(nr, 9).Value = wI.Range("G7").Value
    MsgBox "Invoice " & wI.Range("G4").Value & " saved to Records.", vbInformation
    Exit Sub
E: MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub ApplySettings()
    Application.Calculate
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Invoice")
    Dim d As Long: d = ThisWorkbook.Sheets("Settings").Range("B16").Value
    ws.Range("B46").Value = "1. Payment due within " & d & " days of invoice date."
    MsgBox "Settings applied!", vbInformation
End Sub

Private Function Valid(ws As Worksheet) As Boolean
    Valid = False
    If ws.Range("B10").Value = "Client / Company Name" Or Trim(ws.Range("B10").Value) = "" Then
        MsgBox "Enter a Client Name first.", vbExclamation: ws.Range("B10").Select: Exit Function
    End If
    Dim r As Long
    For r = 18 To 29
        If IsNumeric(ws.Cells(r, 6).Value) Then
            If ws.Cells(r, 6).Value > 0 Then Valid = True: Exit Function
        End If
    Next r
    MsgBox "Enter at least one line item.", vbExclamation: ws.Range("B18").Select
End Function

Private Function SF(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "")
    Next c
    SF = Trim(s)
End Function

'===============================================================================
' LOGO MANAGEMENT
'===============================================================================
Public Sub InsertLogo()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Invoice")
    Dim imgPath As Variant
    imgPath = Application.GetOpenFilename( _
        FileFilter:="Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.gif),*.png;*.jpg;*.jpeg;*.bmp;*.gif", _
        Title:="Select Your Company Logo")
    If imgPath = False Then Exit Sub

    ' Remove existing logo/placeholder
    On Error Resume Next
    ws.Shapes("LogoPlaceholder").Delete
    ws.Shapes("CompanyLogo").Delete
    On Error GoTo 0

    ' Insert image
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture(CStr(imgPath), msoFalse, msoTrue, _
        ws.Range("A2").Left + 2, ws.Range("A2").Top + 2, -1, -1)
    pic.Name = "CompanyLogo"
    pic.OnAction = "RemoveLogo"

    ' Scale to fit header area (max 65x52)
    If pic.Width > 65 Then
        Dim ratio As Double: ratio = 65 / pic.Width
        pic.Width = 65: pic.Height = pic.Height * ratio
    End If
    If pic.Height > 52 Then
        ratio = 52 / pic.Height
        pic.Height = 52: pic.Width = pic.Width * ratio
    End If

    MsgBox "Logo inserted! Click the logo again if you want to remove it.", vbInformation
End Sub

Public Sub RemoveLogo()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Invoice")
    
    If MsgBox("Do you want to remove the current logo?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    On Error Resume Next
    ws.Shapes("CompanyLogo").Delete
    On Error GoTo 0
    ' Restore placeholder
    Dim logoShape As Shape
    Set logoShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("A2").Left + 2, ws.Range("A2").Top + 2, 60, 50)
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
End Sub
