Attribute VB_Name = "modAddin"
Option Explicit

'
' EA_P modAddin
' =============
'
' History:
' 20151010  v1.0.4  nr
'
'
'

Const lg As String = "modAddin > "

'API-Deklarationen
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'Name des Add-ins
Public Const ADDINNAME As String = "EA Makros für Excel"
'Versionsnummer
Public Const VERSION As String = "1.0.4"
'Name des Add-in-Menüs

'Farbkonstanten definieren
Public Const UF_R As Integer = 195
Public Const UF_G As Integer = 218
Public Const UF_B As Integer = 249

'ID des API-Timers
Public lngTimerID As Long

' EA Constanten und Variablen definieren
Public gintWBYear As Integer
Public gstrLastSheet As String
' Public gbInitError As Boolean

Public Const gcstrAIFull As String = "EATools.xlam"
Public Const gcstrAIShort As String = "EATools"
Public Const gcstrKlasse$ = ""
Public Const gcstrTestString$ = "Privat"

Sub Auto_Open()
Dim ll$

    ll = lg & "Auto_Open > "
    Log ll
    
    If Not CheckWBisOK Then
        EA_MenuDeactivate
    Else
        EA_MenuActivate
    End If
    Log ll & "[EOF]"

End Sub

Sub Auto_Close()
    On Error Resume Next
    If ThisWorkbook.Application.AddIns(gcstrAIShort).Installed = False Then
        Call EA_MenuDelete
    End If
End Sub

Public Sub EA_SplashTimer(intSekunden As Integer)
Dim ll$

    ll = lg & "EA_SplashTimer > "
    Log ll
    Log ll & "delay = " & intSekunden
    Log ll & "show splash window"
    frmAddInSplash.Show vbModeless
    'API-Timer aufrufen, Rückgabewert ist ID des Timers.
    'Nach Ablauf der Zeit wird die Prozedur VBA_TimerProc aufgerufen.
    lngTimerID = SetTimer(0, 1, intSekunden * 1000, AddressOf VBA_TimerProc)
    Log ll & "[EOF]"

End Sub

'Schließt Splash-Form Timer-gesteuert
Public Sub VBA_TimerProc(ByVal hWnd As Long, ByVal uint1 As Long, ByVal nEventId As Long, ByVal dwParam As Long)
    Unload frmAddInSplash
    'API-Timer beenden, dabei Timer-ID übergeben
    KillTimer 0, lngTimerID
End Sub

Public Function CheckWBisOK() As Boolean
Dim wb As Workbook
Dim ll$

    On Error GoTo CheckWBisOk_Error
    
    ll = lg & "CheckWBisOK > "
    Log ll
    
    CheckWBisOK = False
    For Each wb In Workbooks
        Log ll & "Workbook found, " & wb.Name
        If Left(LCase(wb.Name), Len(gcstrTestString)) = LCase(gcstrTestString) Then
            Log ll & "OK"   ' & wb.Name
            Log ll & "calling GetWbYear"
            If GetWbYear(wb.Name) = True Then
                Log ll & "GetWbYear=TRUE"
                CheckWBisOK = True
                EA_MenuInsert
                Exit For
            End If
        End If
    Next
    
CheckWBisOk_Error:
    Log ll & CheckWBisOK
    Log ll & "[EOF]"

End Function

Function GetWbYear(Optional strWbName As String) As String
Dim n%, wby$, ll$
Dim intYear

    On Error GoTo Error_GetWbYear_Error
    
    ll = lg & "GetWbYear > "
    Log ll
    Log ll & "passed workbook name=" & strWbName
    
    GetWbYear = ""
    If strWbName = "" Then strWbName = ActiveWorkbook.Name
    Log ll & "aktive workbook=" & strWbName
    
    n = InStr(strWbName, ".xls")
    If n > 0 Then
        wby = Left(strWbName, n - 1)
        wby = Right(wby, 4)
        intYear = CInt(wby)
        If intYear > 1990 Or intYear < 2020 Then
            GetWbYear = intYear
        End If
    End If

Error_GetWbYear_Exit:
    Log ll & "workbook year=" & GetWbYear
    Log ll & "[EOF]"
    Exit Function

Error_GetWbYear_Error:
    GetWbYear = ""
    Resume Error_GetWbYear_Exit

End Function

