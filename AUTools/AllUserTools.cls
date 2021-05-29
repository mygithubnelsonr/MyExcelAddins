VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AllUserTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const lg As String = "Workbook > "

'   Test

Private Sub Workbook_Open()

    Call AUTools_SplashTimer(2)
    Call AUTools_MenuDelete
    Call AUTools_MenuInsert
    
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
                ExportModules
            Case Else
        End Select
    End If
    Log ll & "[EOF]"

End Sub

