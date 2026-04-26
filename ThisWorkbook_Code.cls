' ============================================================
' PASTE THIS CODE INTO THE "ThisWorkbook" CODE MODULE
' (Double-click "ThisWorkbook" in the VBA Project Explorer)
' ============================================================

Private Sub Workbook_Open()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Me.Sheets("Invoice")
    If Not ws Is Nothing Then
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    End If
    On Error GoTo 0
End Sub
