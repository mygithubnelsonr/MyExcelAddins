VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EATools_Workbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const lg As String = "Workbook > "

Private Sub Workbook_Open()
Dim ll$

    On Error GoTo Workbook_Open_Error
    
    ll = lg & "Open > "
    Log ll
    Log ll & "FullName=" & ThisWorkbook.FullName

    EA_SplashTimer 2
    ' EA_MenuDelete
    EA_MenuInsert

'    If CheckWBisOK Then
'        Application.AddIns("EATools").Installed = True
'        Auto_Open
'    End If

Workbook_Open_Exit:
    Log ll & "[EOF]"
    Exit Sub
    
Workbook_Open_Error:
    Log ll & Err.Description
    Resume Workbook_Open_Exit
    
End Sub

Private Sub Workbook_AddinInstall()
Dim ll$

    ll = lg & "AddinInstall > "
    Log ll
'    Log ll & "call CheckWBisOK"
'
'    If CheckWBisOK Then
'        Application.AddIns("EATools").Installed = True
'        Auto_Open
'    End If
    
    Log ll & "[EOF]"

End Sub

Private Sub Workbook_AddinUninstall()
Dim ll$

    ll = lg & "AddinUninstall > "
    Log ll
    
'    Log ll & "call CheckWBisOK"
'
'    If CheckWBisOK Then
'        Application.AddIns("EATools").Installed = True
'        Auto_Open
'    End If
    
    Log ll & "[EOF]"

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim ll$
Dim nRetVal As Long

    ll = lg & "BeforeClose > "
    Log ll

    If Not ThisWorkbook.Saved Then
        nRetVal = MsgBox( _
            "Sollen die Änderungen im Add-In '" & _
            ThisWorkbook.Name & "' gespeichert werden?", _
            vbYesNo + vbQuestion, "Add-In speichern?")
        Select Case nRetVal
            Case vbYes
                ThisWorkbook.IsAddin = True
                ThisWorkbook.Save
            Case Else
        End Select
    End If
    Log ll & "[EOF]"

End Sub

