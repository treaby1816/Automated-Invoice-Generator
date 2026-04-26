VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' --- Step 1: Ensure LicenseData is always hidden ---
    On Error Resume Next
    ThisWorkbook.Sheets("LicenseData").Visible = xlSheetVeryHidden
    On Error GoTo 0

    ' --- Step 2: Validate license ---
    If Not ValidateLicense() Then
        LockWorkbook
        Exit Sub
    End If

    ' --- Step 3: Normal startup ---
    On Error Resume Next
    With ThisWorkbook.Sheets("Invoice")
        .Activate
        ActiveWindow.DisplayGridlines = False
    End With

    Sheets("Dashboard").Activate
    RefreshDashboard
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If SaveAsUI = True Then
        MsgBox "Save As is disabled on this licensed application." & vbNewLine & _
               "Contact your vendor if you need to migrate this file.", _
               vbExclamation, "Action Restricted"
        Cancel = True
    End If
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ThisWorkbook.Sheets("LicenseData").Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub
