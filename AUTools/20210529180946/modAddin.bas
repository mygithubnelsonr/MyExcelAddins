Attribute VB_Name = "modAddin"
Option Explicit

Const lg As String = "modAddin > "
'Name des Add-ins
Public Const ADDINNAME As String = "AUTools"
Public Const gcstrAddin$ = "AUTools.xlam"
'Versionsnummer
Public Const gcstrVersion As String = "1.1.0"
'Name des Add-in-Menüs
Public Const gcstrMenuName As String = "All User Tools"
'Farbkonstanten definieren
Public Const UF_R As Integer = 195
Public Const UF_G As Integer = 218
Public Const UF_B As Integer = 249


Public Sub AUTools_SplashTimer(intSeconds As Integer)

    On Error Resume Next
    frmAddinSplash.Show vbModeless
    'API-Timer aufrufen, Rückgabewert ist ID des Timers.
    'Nach Ablauf der Zeit wird die Prozedur VBA_TimerProc aufgerufen.
    lngTimerID = SetTimer(0, 1, intSeconds * 1000, AddressOf VBA_Timer1)
    
End Sub

'Schließt Splash-Form Timer-gesteuert
Public Sub VBA_Timer1(ByVal hWnd As Long, ByVal uint1 As Long, ByVal nEventId As Long, ByVal dwParam As Long)

    On Error Resume Next
    Unload frmAddinSplash
    ' API-Timer beenden, dabei Timer-ID übergeben
    KillTimer 0, lngTimerID
    
End Sub

