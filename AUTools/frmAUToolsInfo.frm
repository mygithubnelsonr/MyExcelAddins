VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAUToolsInfo 
   Caption         =   "All User Tools"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "frmAUToolsInfo.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAUToolsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const lg As String = "frmAUToolsInfo > "


Private Sub cmd001_Click()
    Call CopyUpperCell
End Sub

Private Sub cmd002_Click()
    Call ConvertToUpper
End Sub


Private Sub cmd003_Click()
    ConvertToCapitals
End Sub

Private Sub UserForm_Initialize()
Dim ll$
Const intRand% = 16

    ll = lg & "Initialize > "
    Log ll
    With Me
        .Image1.Width = .Width
        .lblHeader.Caption = Me.Caption
        .lblHeader.Left = 0
        .lblHeader.Width = Me.Width
        .lblVersion = gcstrVersion
        .lblClose.Left = .Width - .lblClose.Width
        .MultiPage1.Width = .Width - intRand
        .MultiPage1.Value = 0
    End With

End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.Font.Size = 10
End Sub

Private Sub lblClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.Font.Size = 11
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub ListBoxFiles_Change()
Dim n%, C%

    C = 0
    For n = 0 To ListBoxFiles.ListCount - 1
        If ListBoxFiles.Selected(n) Then C = C + 1
    Next
    If C = 0 Then
        cmdStartJob.Enabled = False
    Else
        cmdStartJob.Enabled = True
    End If
End Sub

Private Sub chkSelectAll_Click()
Dim blnSel As Boolean
Dim n%
Dim Entry

    With Me
        If .chkSelectAll Then
            For n = 0 To ListBoxFiles.ListCount - 1
                ListBoxFiles.Selected(n) = True
            Next
            .chkSelectAll.Caption = "unselect all"
        Else
            For n = 0 To ListBoxFiles.ListCount - 1
                ListBoxFiles.Selected(n) = False
            Next
            .chkSelectAll.Caption = "select all"
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

