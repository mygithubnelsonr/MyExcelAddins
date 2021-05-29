Attribute VB_Name = "modMain"
Option Explicit

'
' EA_P modMain
' ============
'

Const lg As String = "modMain > "

Sub Log(strMsg$)
    Debug.Print "[" & Now() & "] " & gcstrAIShort & " > " & strMsg
End Sub

Function GetWbShortName() As String
Dim strYear$

    On Error GoTo GetWbShortName_Error

    ' strYear = GetWbYear(ActiveWorkbook.Name)
    
    GetWbShortName = Left(ThisWorkbook.Name, InStr(ThisWorkbook.Name, ".") - 1)
    
GetWbShortName_Exit:
    Exit Function
    
GetWbShortName_Error:
    Resume GetWbShortName_Exit

End Function


Function GetFirstOfMonth(intMonth As Integer) As Currency
Dim ws As Worksheet
Dim rng As Range
Dim Element
Dim n&, R&, C&
Dim strResult As Currency

    On Error GoTo GetFirstOfMonth_Error
    
    Set ws = ActiveWorkbook.Worksheets("Bank")
    Set rng = ws.Range("A80:A1400")
    
    R = rng.Rows.Count
    
    For Each Element In rng
        If Month(Element) = intMonth Then
            R = Element.row
            strResult = ws.Cells(R, 6)
            Exit For
        Else
            strResult = 0
        End If
        If n > 2000 Then Exit For
    Next

GetFirstOfMonth_Exit:
    GetFirstOfMonth = strResult
    Exit Function
    
GetFirstOfMonth_Error:
    strResult = 0
    Resume GetFirstOfMonth_Exit

End Function

Function ValidatePath(strPath)
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    ValidatePath = strPath
End Function

Public Sub ExportModules()
Dim wb As Workbook

    Set wb = ThisWorkbook
    EAToolsExportModules wb

End Sub

Public Sub EAToolsExportModules(wb As Workbook)
Dim blnExport As Boolean
Dim strPath$, strFileName$, strExportPath As String
Dim cmpComponent As VBIDE.VBComponent

    On Error GoTo EAToolsExportModules_Error
    
    strPath = EAToolsCreateExportFolder(wb)
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
        strExportPath, vbInformation, "EAToolsExportModules: Done"

EAToolsExportModules_Exit:
    Exit Sub

EAToolsExportModules_Error:
    MsgBox Err.Description, vbCritical, "EAToolsExportModules: Error"
    Resume EAToolsExportModules_Exit

End Sub

Function EAToolsCreateExportFolder(wb As Workbook) As String
Dim WshShell As Object
Dim oFso As Object
Dim strExportAdd$, strWbName$, strExportPath$, Path$, strWbPath As String
Dim ar, Element

    On Error GoTo EAToolsCreateExportFolder_Error
    
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
        EAToolsCreateExportFolder = strExportPath
    Else
        EAToolsCreateExportFolder = "Error"
    End If
    
EAToolsCreateExportFolder_Exit:
    Exit Function

EAToolsCreateExportFolder_Error:
    EAToolsCreateExportFolder = "Error"
    Resume EAToolsCreateExportFolder_Exit

End Function

