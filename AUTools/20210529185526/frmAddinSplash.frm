VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddinSplash 
   Caption         =   "frmAddInSplash"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   OleObjectBlob   =   "frmAddinSplash.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAddinSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub UserForm_Initialize()
    With Me
        .Caption = "Info"
        .lblTitel = gcstrMenuName
        .lblTitel.ForeColor = vbWhite
        .lblVersion = "Version " & gcstrVersion
        .lblVersion.ForeColor = vbWhite
        .lblCopyright = "Copyright © 2000 - " & Year(Now) & vbLf & "NRSoft Robert Nelson"
        .lblCopyright.ForeColor = vbWhite
        'Wenn Version der Office-Anwendung 11 oder darüber, dann...
        If Val(Application.Version) > 10 Then
            '... Userform-Hintergrund auf typisches Baby-Blau von Office
            '2003 einstellen. Dazu muss die Backstyle-Eigenschaft von Labels
            'und anderen "durchscheinenden" Userform-Controls auf 0 (Transparent)
            'eingestellt sein
            '.BackColor = RGB(UF_R, UF_G, UF_B)
            '.BackColor = RGB(30, 42, 36)
        End If
    End With
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Unload Me
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub
