VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddInSplash 
   Caption         =   "frmAddInSplash"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   OleObjectBlob   =   "frmAddInSplash.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAddInSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'© 2006, Ralf Nebelo

Private Sub UserForm_Initialize()
    With Me
        'Dialogtitel festlegen
        .Caption = "Info"
        'Add-in-Name, Version und Copyright-Text festlegen
        .lblTitel = modAddin.ADDINNAME
        .lblVersion = "Version " & modAddin.VERSION
        .lblCopyright = "Copyright © 2006-2014 NRSoft"
        
        'Wenn Version der Office-Anwendung 11 oder darüber, dann...
        If Val(Application.VERSION) > 10 Then
            '... Userform-Hintergrund auf typisches Baby-Blau von Office
            '2003 einstellen. Dazu muss die Backstyle-Eigenschaft von Labels
            'und anderen "durchscheinenden" Userform-Controls auf 0 (Transparent)
            'eingestellt sein
            .BackColor = RGB(modAddin.UF_R, modAddin.UF_G, modAddin.UF_B)
        End If
    End With
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub
