Attribute VB_Name = "Module2"
Option Explicit

' ================================================================
' DEVELOPER DEPLOYMENT CHECKLIST
' ================================================================
' 1. Change salt string in HashFingerprint() to your own secret value
' 2. Open workbook on CLIENT's machine
' 3. Run GenerateLicenseForClient() — copy the output key
' 4. Paste license key into LicenseData!B1
' 5. Fill client company name in LicenseData!B2
' 6. Set expiry date in LicenseData!B3 (format: DD-MMM-YYYY)
' 7. Fill Settings!B2:B5 with client's company details
' 8. Lock VBA project: VBA Editor > Tools > VBAProject Properties > Protection
' 9. Set VBA project password (keep this password — do NOT give to client)
' 10. Save and close workbook
' 11. Reopen to confirm license validation passes
' 12. Deliver workbook to client
' ================================================================
' RENEWAL PROCESS (remote — no site visit needed):
' 1. Client sends you their machine fingerprint (run GenerateLicense)
' 2. You generate new key + new expiry date
' 3. Send key string to client via email/WhatsApp
' 4. Client runs RenewLicense macro and pastes key + new date
' ================================================================

Public Function GetMachineFingerprint() As String
    Dim strMAC As String
    Dim strVolSerial As String
    Dim strPCName As String

    ' --- Get MAC Address via WMI ---
    On Error Resume Next
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery( _
        "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each objItem In colItems
        strMAC = objItem.MACAddress
        Exit For
    Next
    On Error GoTo 0

    ' --- Get Hard Drive Volume Serial Number ---
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    strVolSerial = CStr(fso.GetDrive("C:\").SerialNumber)
    On Error GoTo 0

    ' --- Get PC/Computer Name ---
    strPCName = Environ("COMPUTERNAME")

    ' --- Combine into fingerprint ---
    GetMachineFingerprint = strMAC & "|" & strVolSerial & "|" & strPCName
End Function

Public Function HashFingerprint(fingerprint As String) As String
    Dim salt As String
    salt = "MORESTA_SECRET_2026_X99"  ' <-- DEVELOPER: change this to your own secret

    Dim combined As String
    combined = fingerprint & salt

    Dim result As Long
    Dim i As Integer
    result = 0
    For i = 1 To Len(combined)
        result = result Xor (Asc(Mid(combined, i, 1)) * i)
    Next i

    ' Format as a readable license key
    Dim raw As String
    raw = Hex(Abs(result))
    Do While Len(raw) < 12
        raw = "0" & raw
    Loop

    HashFingerprint = Left(raw, 4) & "-" & Mid(raw, 5, 4) & "-" & Right(raw, 4)
End Function

Public Sub GenerateLicenseForClient()
    Dim fp As String
    Dim key As String
    fp = GetMachineFingerprint()
    key = HashFingerprint(fp)

    MsgBox "=== LICENSE GENERATION TOOL ===" & vbNewLine & vbNewLine & _
           "Machine Fingerprint:" & vbNewLine & fp & vbNewLine & vbNewLine & _
           "Generated License Key:" & vbNewLine & key & vbNewLine & vbNewLine & _
           "Paste this key into LicenseData!B1" & vbNewLine & _
           "Set expiry date in LicenseData!B3", _
           vbInformation, "License Generator"
End Sub

Public Function ValidateLicense() As Boolean
    On Error GoTo LicenseFail

    Dim wsLic As Worksheet
    Set wsLic = ThisWorkbook.Sheets("LicenseData")

    ' --- Read stored values ---
    Dim storedKey As String
    Dim expiryDate As Date
    Dim registeredCompany As String
    storedKey = Trim(wsLic.Range("B1").Value)
    registeredCompany = Trim(wsLic.Range("B2").Value)

    ' --- Check expiry ---
    If wsLic.Range("B3").Value = "" Then GoTo LicenseFail
    expiryDate = CDate(wsLic.Range("B3").Value)
    If Date > expiryDate Then
        MsgBox "Your license expired on " & Format(expiryDate, "DD-MMM-YYYY") & _
               "." & vbNewLine & "Please contact your vendor to renew.", _
               vbCritical, "License Expired"
        ValidateLicense = False
        Exit Function
    End If

    ' --- Check hardware fingerprint ---
    Dim currentFP As String
    Dim expectedKey As String
    currentFP = GetMachineFingerprint()
    expectedKey = HashFingerprint(currentFP)

    If storedKey <> expectedKey Then
        ' --- Grace login handling ---
        Dim graceLeft As Integer
        graceLeft = wsLic.Range("B6").Value
        If graceLeft > 0 Then
            wsLic.Range("B6").Value = graceLeft - 1
            MsgBox "Warning: License mismatch detected." & vbNewLine & _
                   "Grace logins remaining: " & (graceLeft - 1) & vbNewLine & _
                   "Contact your vendor immediately.", vbExclamation, "License Warning"
            ValidateLicense = True  ' Allow but warn
        Else
            ValidateLicense = False
        End If
        Exit Function
    End If

    ' --- Store activation info on first valid open ---
    If wsLic.Range("B4").Value = "" Then
        wsLic.Range("B4").Value = Date
        wsLic.Range("B5").Value = currentFP
    End If

    ValidateLicense = True
    Exit Function

LicenseFail:
    ValidateLicense = False
End Function

Public Sub LockWorkbook()
    Dim ws As Worksheet

    ' --- Ensure UNAUTHORIZED sheet exists, if not create a temporary one ---
    Dim unauthExists As Boolean: unauthExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "UNAUTHORIZED" Then unauthExists = True: Exit For
    Next ws
    
    If Not unauthExists Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "UNAUTHORIZED"
        ws.Range("A1").Value = "ACCESS DENIED - UNLICENSED MACHINE"
        ws.Range("A1").Font.Size = 24
        ws.Range("A1").Font.Color = vbRed
        ws.Range("A1").Font.Bold = True
    End If

    ' --- Hide all sheets ---
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "UNAUTHORIZED" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' --- Show lockout screen ---
    ThisWorkbook.Sheets("UNAUTHORIZED").Visible = xlSheetVisible
    ThisWorkbook.Sheets("UNAUTHORIZED").Activate

    MsgBox "This application is not licensed for this machine." & vbNewLine & _
           "Contact your vendor to activate or renew your license.", _
           vbCritical, "Access Denied"

    ' --- Auto-close after 15 seconds ---
    Application.OnTime Now + TimeValue("00:00:15"), "AutoClose"
End Sub

Public Sub AutoClose()
    ThisWorkbook.Close SaveChanges:=False
End Sub

Public Sub RenewLicense()
    Dim newKey As String
    Dim newExpiry As String

    newKey = InputBox("Enter your new License Key provided by your vendor:", "License Renewal")
    If newKey = "" Then Exit Sub

    newExpiry = InputBox("Enter new Expiry Date (DD-MMM-YYYY):", "License Renewal")
    If newExpiry = "" Then Exit Sub

    On Error GoTo BadDate
    Dim testDate As Date
    testDate = CDate(newExpiry)

    ' --- Validate the new key ---
    Dim fp As String
    fp = GetMachineFingerprint()
    Dim expectedKey As String
    expectedKey = HashFingerprint(fp)

    If Trim(newKey) = expectedKey Then
        Dim wsLic As Worksheet
        Set wsLic = ThisWorkbook.Sheets("LicenseData")
        wsLic.Range("B1").Value = newKey
        wsLic.Range("B3").Value = testDate
        wsLic.Range("B6").Value = 3  ' Reset grace logins
        MsgBox "License renewed successfully! Valid until " & _
               Format(testDate, "DD-MMM-YYYY"), vbInformation, "Renewal Successful"
    Else
        MsgBox "Invalid license key. Please contact your vendor.", vbCritical, "Invalid Key"
    End If
    Exit Sub

BadDate:
    MsgBox "Invalid date format. Please use DD-MMM-YYYY (e.g., 31-Dec-2025).", vbExclamation
End Sub
