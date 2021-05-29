Attribute VB_Name = "modAPI"
Option Explicit

Const lg As String = "modAPI > "

'API-Deklarationen

'ID des API-Timers
Public lngTimerID As Long
'
' Timer Functions
'
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'
' File download
'
Public Declare Function URLDownloadToFile Lib "urlmon.dll" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
        ByVal szURL As String, ByVal szFileName As String, _
        ByVal Reserved As Long, ByVal fnCB As Long) As Long

