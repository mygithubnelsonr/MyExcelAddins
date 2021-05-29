Attribute VB_Name = "modMain"
Option Explicit


Public Const lg As String = "modMain > "
Public gstrStartPath$, gstrSourcePath$, gstrDatabase$, gstrCurrUser$


Sub Split2Cells()
Dim strText$
Dim ar
Dim R%, C%, n%

    strText = ActiveCell.Value2
    ar = Split(strText, " ")
    
    R = ActiveCell.Row
    C = ActiveCell.Column
    
    If UBound(ar) > 1 Then
        For n = 0 To UBound(ar)
            If Trim(ar(n)) <> "" Then
                Debug.Print ar(n)
                Cells(R, C + 1).Value = ar(n)
                C = C + 1
            End If
        Next
    End If
    
End Sub

Sub CopyUpperCell()
Attribute CopyUpperCell.VB_ProcData.VB_Invoke_Func = "K\n14"

' Tastenkombination: Shift-Strg+k
'

    On Error GoTo CopyUpperCell_Error
    
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:=xlFillCopy
    ActiveCell.Offset(1, 1).Range("A1").Select
    
CopyUpperCell_Exit:
    Exit Sub

CopyUpperCell_Error:
    Resume CopyUpperCell_Exit
    
End Sub

Sub InsertDateTime()
    ActiveCell.FormulaR1C1 = "=NOW()"
    ActiveCell.Select
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Sub ConvertToUpper()
Attribute ConvertToUpper.VB_ProcData.VB_Invoke_Func = "U\n14"
'
' Tastenkombination: Shift-Strg+k
'
Dim rng As Range
Dim Element As Range

    On Error GoTo ConvertToUpper_Error
    
    Set rng = Selection
    
    For Each Element In rng
        Element = UCase(Element)
    Next

ConvertToUpper_Exit:
    Exit Sub
    
ConvertToUpper_Error:
    Resume ConvertToUpper_Exit

End Sub

Sub ProperCase()
Dim l, NewString, i
Dim ar
Dim rngR As Range
Dim R, C

    Set rngR = Selection

    For Each C In rngR
        l = Len(C)
        If l > 0 Then
            ar = Split(LCase(C), " ")
            NewString = ""
            For i = 0 To UBound(ar)
                ar(i) = UCase(Left(ar(i), 1)) & Mid(ar(i), 2, l)
                NewString = NewString & ar(i) & " "
            Next
            C.Value = RTrim(NewString)
        End If
    Next
End Sub

Sub ConvertToCapitals()
Attribute ConvertToCapitals.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' Tastenkombination: Shift-Strg+c
'
Dim ar
Dim n%, t$, strNew$, strOld$
Dim rng As Range
Dim Element As Range

    On Error GoTo ConvertToCapitals_Error
    
    Set rng = Selection
    For Each Element In rng
        strOld = Trim(ActiveCell)
        strNew = ""
        ar = Split(LCase(strOld), " ")
        For n = 0 To UBound(ar)
            t = ar(n)
            If Trim(t) <> "" Then
                t = UCase(Left(t, 1)) & Right(t, Len(t) - 1)
                If n = UBound(ar) Then
                    strNew = strNew & t
                Else
                    strNew = strNew & t & " "
                End If
            End If
        Next
        ar = Split(Trim(strNew), ".")
        strNew = ""
        For n = 0 To UBound(ar)
            t = ar(n)
            t = UCase(Left(t, 1)) & Right(t, Len(t) - 1)
            If n = UBound(ar) Then
                strNew = strNew & t
            Else
                strNew = strNew & t & "."
            End If
        Next
        ar = Split(Trim(strNew), "-")
        strNew = ""
        For n = 0 To UBound(ar)
            t = ar(n)
            t = UCase(Left(t, 1)) & Right(t, Len(t) - 1)
            If n = UBound(ar) Then
                strNew = strNew & t
            Else
                strNew = strNew & t & "-"
            End If
        Next
        ar = Split(Trim(strNew), "(")
        strNew = ""
        For n = 0 To UBound(ar)
            t = ar(n)
            t = UCase(Left(t, 1)) & Right(t, Len(t) - 1)
            If n = UBound(ar) Then
                strNew = strNew & t
            Else
                strNew = strNew & t & "("
            End If
        Next
        ar = Split(Trim(strNew), ",")
        strNew = ""
        For n = 0 To UBound(ar)
            t = ar(n)
            t = UCase(Left(t, 1)) & Right(t, Len(t) - 1)
            If n = UBound(ar) Then
                strNew = strNew & t
            Else
                strNew = strNew & t & ","
            End If
        Next
        Element = Trim(strNew)
    Next
    
    ' ActiveCell.Offset(1, 0).Range("A1").Select
    
ConvertToCapitals_Exit:
    Exit Sub

ConvertToCapitals_Error:
    Resume ConvertToCapitals_Exit
    
End Sub

Sub Log(strMsg$)
'Dim oFso As New FileSystemObject
'Dim ofil As File
'Dim olog As TextStream
'Dim Result
'Dim strLogfile$
'
'    ' On Error Resume Next
'    strLogfile = ThisWorkbook.Path & "\OMWebAI.log"
'    If Not oFso.FileExists(strLogfile) Then
'        oFso.CreateTextFile (strLogfile)
'    End If
'
'    Set ofil = oFso.GetFile(strLogfile)
'    Set olog = ofil.OpenAsTextStream(ForAppending)
'    olog.WriteLine "[" & Now() & "] [" & gstrCurrUser & "] OMWebAI > " & strMsg
'    olog.Close
'    Set ofil = Nothing
'    Set oFso = Nothing
    
End Sub

Function GetFormatedDate()
    GetFormatedDate = Format(Now(), "yyyymmdd")
End Function

Function GetUserName() As String
Dim objWSHNetwork As Object
    On Error Resume Next

    Set objWSHNetwork = CreateObject("WScript.Network")
    GetUserName = objWSHNetwork.UserName
    Set objWSHNetwork = Nothing

End Function

Sub GUIShow_AUToolsInfo()
    On Error Resume Next
    frmAUToolsInfo.Show vbModal
End Sub

Sub GUIShow_AUToolsAbout()
    ' Call OMWebAIAddInSplashTimer(20)
    frmAddinSplash.Show vbModeless
End Sub

Public Sub AddTimer2(intMilliSeconds As Integer)
    On Error GoTo AddTimer2_Error
    'API-Timer aufrufen, Rückgabewert ist ID des Timers.
    'Nach Ablauf der Zeit wird die Prozedur VBA_TimerProc aufgerufen.
    'frmOMWebTools!lblAction.BackColor = vbRed
    lngTimerID = SetTimer(0, 1, intMilliSeconds, AddressOf VBA_Timer2)
    Exit Sub
    
AddTimer2_Error:
    
End Sub

Public Sub VBA_Timer2(ByVal hWnd As Long, ByVal uint1 As Long, ByVal nEventId As Long, ByVal dwParam As Long)
Static n%
Dim s$, ll$

    ll = lg & "VBA_Timer2 > "
    On Error GoTo VBA_Timer2_Error
    Select Case n
        Case 1: s = "|"
        Case 2: s = "/"
        Case 3: s = "--"
        Case 4
            s = "\"
            n = 0
    End Select
    Log s
    n = n + 1
    
VBA_Timer2_Exit:
    Exit Sub
    
VBA_Timer2_Error:
    Resume VBA_Timer2_Exit
End Sub

Public Sub KillTimer2()
    KillTimer 0, lngTimerID
End Sub

Function ValidatePath(strPath)
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    ValidatePath = strPath
End Function

Public Sub AUTExportModules(wb As Workbook)
Dim blnExport As Boolean
Dim strPath$, strFileName$, strExportPath As String
Dim cmpComponent As VBIDE.VBComponent

    On Error GoTo AUTExportModules_Error
    
    strPath = AUTCreateExportFolder(wb)
    ' Debug.Print strPath
    If strPath = "Error" Then
        MsgBox "Export Folder not exist", vbCritical, "AUTExportModules: Error"
        Exit Sub
    End If
    
    ' NOTE: This workbook must be open in Excel.
    ' the passed workbook wb must be open
    If wb.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & vbLf & _
            "not possible to export the code", vbCritical, "AUTExportModules: Error"
        Exit Sub
    End If
    
    strExportPath = ValidatePath(strPath)
    For Each cmpComponent In wb.VBProject.VBComponents
        blnExport = True
        strFileName = cmpComponent.Name
        ' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                strFileName = strFileName & ".cls"
            Case vbext_ct_MSForm
                strFileName = strFileName & ".frm"
            Case vbext_ct_StdModule
                strFileName = strFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                strFileName = strFileName & ".cls"
                ' blnExport = False
        End Select
        
        If blnExport Then
            ' Export the component to a text file.
            cmpComponent.Export strExportPath & strFileName
            ' remove it from the project if you want
            ' wkbSource.VBProject.VBComponents.Remove cmpComponent
        End If
    Next cmpComponent
    MsgBox "All modules are successfully exported to " & vbLf & _
        strExportPath, vbInformation, "AUTExportModules: Done"

AUTExportModules_Exit:
    Exit Sub

AUTExportModules_Error:
    MsgBox Err.Description, vbCritical, "AUTExportModules: Error"
    Resume AUTExportModules_Exit

End Sub

Function AUTCreateExportFolder(wb As Workbook) As String
Dim WshShell As Object
Dim oFso As Object
Dim strExportAdd$, strWbName$, strExportPath$, Path$, strWbPath As String
Dim ar, Element

    On Error GoTo AUTCreateExportFolder_Error
    
    Set WshShell = CreateObject("WScript.Shell")
    Set oFso = CreateObject("scripting.filesystemobject")
    
    strWbPath = ValidatePath(wb.Path)
    strWbName = Left(wb.Name, InStr(wb.Name, ".") - 1)
    strExportPath = strWbPath & "Code\" & strWbName & "\" & Format(Now(), "yyyymmddhhnnss")
    strExportAdd = Replace(strExportPath, strWbPath, "")
    
    ar = Split(strExportAdd, "\")
    Path = strWbPath
    For Each Element In ar
        Path = ValidatePath(Path & Element)
        If (Dir(Path, vbDirectory) = "") Then
          MkDir Path
        End If
    Next
    
    If oFso.FolderExists(strExportPath) = True Then
        AUTCreateExportFolder = strExportPath
    Else
        AUTCreateExportFolder = "Error"
    End If
    
AUTCreateExportFolder_Exit:
    Exit Function

AUTCreateExportFolder_Error:
    AUTCreateExportFolder = "Error"
    Resume AUTCreateExportFolder_Exit

End Function

Sub ExportModules()
Dim wb As Workbook

    Set wb = ThisWorkbook
    AUTExportModules wb
End Sub

