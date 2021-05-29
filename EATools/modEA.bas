Attribute VB_Name = "modEA"
Option Explicit

'
' EA modEA
' ========
'

Const lg As String = "modEA > "

Public Const gcintMwst1% = 19
Public Const gcintMwst2% = 7
' Public Const gcstrBankKonten As String = "20273,P,22535,G"

Sub EA_About()
    'Userform frmAddInSplash nicht modal(!) anzeigen
    frmAddInSplash.Show vbModeless
End Sub

Sub EA_ResetKey()
    Application.OnKey "{RETURN}"
    Application.OnKey "{ENTER}"
End Sub

Sub EA_NewTemplate()
Dim sWorkBook$, sPath$, sYear$
    
    On Error GoTo EA_NewTemplate_Error
    
    Sheets("Welcome").Activate
    sPath = ValidatePath(ActiveWorkbook.Path)
    sWorkBook = sPath & gcstrTestString & "Template" & ".xlsm"
    ActiveWorkbook.SaveAs Filename:=sWorkBook, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    EA_CleanUpBank "KP-20273"
    EA_CleanUpBank "KP-22535"
    ' EA_CleanUpKasse
    EA_CleanUpJornal
    EA_CleanUpVerlauf
    EA_CleanUpEA
    
    Sheets("Welcome").Activate

EA_NewTemplate_Exit:
    Exit Sub
    
EA_NewTemplate_Error:
    Resume EA_NewTemplate_Exit:
    
End Sub

Sub EA_CleanUpEA()
Dim rng As Range
Dim iFirstRow%, iLastRow%
    
    Sheets("Anschaffungen").Activate
    
    Range("2:19").Select
    Selection.ClearContents
    Selection.ClearComments
    Cells.Select
    Selection.Columns.AutoFit

    Range("A2").Select
    
End Sub

Sub EA_CleanUpVerlauf()
Dim rng As Range
Dim iFirstRow%, iLastRow%
    
    Sheets("Verlauf").Activate
    
    Range("D4:D16").Select
    Selection.ClearContents
    Selection.ClearComments

    Range("A1").Select
End Sub

Sub EA_CleanUpJornal()
Dim rng As Range
Dim iFirstRow%, iLastRow%
    
    Sheets("Jornal").Activate
    iFirstRow = Range("JPFirstRow").row
    iLastRow = Range("JPLastRow").Value
    
    Range(CStr(iFirstRow) & ":" & CStr(iLastRow)).Select
    Selection.ClearContents
    Range(CStr(iFirstRow + 1) & ":" & CStr(iLastRow)).Select
    Selection.Delete

    Application.GoTo Reference:="JPFirstRow"
    
End Sub

Sub EA_CleanUpKasse()
Dim rng As Range
Dim iFirstRow%, iLastRow%
    
    Sheets("Kasse").Activate
    iFirstRow = Range("KPFirstRow").row
    iLastRow = Range("KPLastRow").Value
    
    EA_SingleTexteK
    
    Range("C2").Value = 0
    
    Range("A" & CStr(iFirstRow) & ":F" & CStr(iLastRow)).Select
    Selection.ClearContents

    Range("A" & CStr(iFirstRow) & ":E" & CStr(iLastRow)).Select
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Style = "Normal"

    Range("B3:C3").Select
    Selection.Copy
    
    Range("B" & CStr(iFirstRow) & ":C" & CStr(iLastRow)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Application.GoTo Reference:="KPFirstRow"
    
End Sub

Sub EA_SingleTexteK()
Dim col As New Collection
Dim rng As Range
Dim iFirstRow%, iLastRow%, iCount%, n%
Dim Element, CV

    Application.DisplayAlerts = False
    
    Sheets("Kasse").Activate
    
    iFirstRow = Range("KPFirstRow").row
    iLastRow = Range("KPLastRow").Value
    
    Range("E" & CStr(iFirstRow) & ":E" & CStr(iLastRow)).Select
    Set rng = Selection
    
    On Error Resume Next
    For Each Element In rng
        CV = CStr(Element.Value)
        col.Add Element.Value2, Element.Value2
    Next
    On Error GoTo 0
    iCount = col.Count
    
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    ' vorhandene Einträge erst löschen
    Range("KPTexte").ClearContents
    ' mit neuen Werten befüllen
    Range("E4").Select
    
    For Each Element In col
        ActiveCell.Value = Element
        ActiveCell.Offset(1, 0).Select
    Next
    
    For n = ActiveCell.row To iFirstRow - 1
        ActiveCell.Value = "zz"
        ActiveCell.Offset(1, 0).Select
    Next
    
    Application.GoTo Reference:="KPTexte"
    Selection.Sort Key1:=Range("E4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    Application.GoTo Reference:="KPFirstRow"
    
End Sub

Sub EA_CleanUpBank(strKonto As String)
Dim rng As Range
Dim iFirstRow%, iLastRow%
    
    Sheets(strKonto).Activate
    iFirstRow = Range("BPFirstRow").row
    iLastRow = Range("BPLastRow").Value
    
    EA_SingleTexteB strKonto
    
    Range("E2").Value = "Übertrag Jahr 20xx"
    Range("F2").Value = 0
    
    Range("A" & CStr(iFirstRow) & ":G" & CStr(iLastRow)).Select
    Selection.ClearContents
    
    Range("A" & CStr(iFirstRow) & ":F" & CStr(iLastRow)).Select
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Style = "Normal"

    Range("C3:F3").Select
    Selection.Copy
    Range("C" & CStr(iFirstRow) & ":F" & CStr(iLastRow)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Application.GoTo Reference:="BPFirstRow"

End Sub

Sub EA_SingleTexteB(sKOnto As String)
Dim col As New Collection
Dim rng As Range
Dim iFirstRow%, iLastRow%, iCount%, n%
Dim Element, CV

    Application.DisplayAlerts = False
    
    Sheets(sKOnto).Activate
    
    iFirstRow = Range("BPFirstRow").row
    iLastRow = Range("BPLastRow").Value
    
    Range("B" & CStr(iFirstRow) & ":B" & CStr(iLastRow)).Select
    Set rng = Selection
    
    On Error Resume Next
    For Each Element In rng
        CV = CStr(Element.Value)
        col.Add Element.Value2, Element.Value2
    Next
    On Error GoTo 0
    iCount = col.Count
    
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    ' vorhandene Einträge erst löschen
    Range("BPTexte").ClearContents
    ' mit neuen Werten befüllen
    Range("B" & Range("BPTexte").row).Select
    
    For Each Element In col
        ActiveCell.Value = Element
        ActiveCell.Offset(1, 0).Select
    Next
    
    For n = ActiveCell.row To iFirstRow - 1
        ActiveCell.Value = "zz"
        ActiveCell.Offset(1, 0).Select
    Next
    
    Application.GoTo Reference:="BPTexte"
    Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    Application.GoTo Reference:="BPFirstRow"
    
End Sub

Sub EA_Bank2Jornal()
Dim rngT As Range, rngLines As Range, rngJ As Range
Dim intJRow%, intKRK As Integer
Dim R&, C As Long
Dim ll$, strKShort$, strActiveSheet As String
Dim Kommevon, BankDatum, BankText, BankEin, BankAus, Bankbetrag, konto, result As Variant

    On Error GoTo EA_Bank2Jornal_Error
    
    ll = lg & "EA_Bank2Jornal > "
    Log ll
    
    strActiveSheet = Application.ActiveSheet.Name
    
    If Not IsKonto(strActiveSheet) Then
        Exit Sub
    End If
    
    strKShort = GetKontoShort(strActiveSheet)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Set rngLines = Selection
    For Each rngT In rngLines
        Cells(rngT.row, 1).Select
        If Selection.Style <> "Gut" Then
            BankDatum = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            BankText = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            intKRK = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            BankEin = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            BankAus = (ActiveCell() * -1)
            
            If BankEin > 0 Then
                Bankbetrag = BankEin
            Else
                Bankbetrag = BankAus
            End If
            
            If IsNumeric(intKRK) Then
                ' check and create new account
                result = JornalKontoExist(intKRK)
                If result = False Then
                    MsgBox ("Konto " + CStr(intKRK) + " nicht gefunden!")
                    Exit Sub
                End If
                
                C = result  ' column with konto
                ' nächste freie zeile ist in Range("JLastEntry") hinterlegt
                intJRow = Sheets("Jornal").Range("JPLastRow").Value + 1
                Set rngJ = Sheets("Jornal").Rows(intJRow)
                rngJ.Cells(1, 1).Value = BankDatum
                rngJ.Cells(1, 1).NumberFormat = "DD.MM."
                rngJ.Cells(1, 1).HorizontalAlignment = xlCenter
                rngJ.Cells(1, 2).Value = strKShort
                rngJ.Cells(1, 3).Value = BankText
                rngJ.Cells(1, C).Value = Bankbetrag
                
                Set rngJ = Sheets("Jornal").Rows(intJRow + 1)
                rngJ.EntireRow.Insert
                ActiveCell.Offset(0, -4).Range("A1").Select
                ' bearbeitete zeile farblich markieren
                Selection.Style = "Gut"
                ActiveCell.Offset(1, 0).Range("A1").Select
            End If
        End If
    Next
    Application.ScreenUpdating = True

EA_Bank2Jornal_Exit:
    Log ll & "[EOF]"
    Exit Sub
    
EA_Bank2Jornal_Error:
    Log ll & Err.Description
    Resume EA_Bank2Jornal_Exit
    
End Sub

Sub EA_Kasse2Jornal()
Const intStoreCell As Integer = 2
Dim rngCKonto As Range, rngLines As Range, rngT As Range
Dim intJRow%, intBankkonto As Integer
Dim blnFound As Boolean
Dim datBankDatum As Date
Dim ll$, strKShort$, strActiveSheet$, strBankKonto$, strBuchungsText As String
Dim varAusgaben, varBestand, varEinnahmen, C, result As Variant

    On Error GoTo EA_Kasse2Jornal_Error
    
    ll = lg & "EA_Kasse2Jornal > "
    Log ll
    
    strActiveSheet = Application.ActiveSheet.Name
    
    If Not strActiveSheet = "Kasse" Then
        Err.Raise 900, "EA_Kasse2Jornal", "Sheet is a bankaccount! Exit programm."
    End If
    
    strKShort = "K"
    Set rngLines = Selection

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    For Each rngT In rngLines
        Cells(rngT.row, 1).Select   ' column ausgaben
        Debug.Print "Cell Style=" & Selection.Style
        ' wenn schon bearbeitet dann nächster eintrag
        If Not Selection.Style = "Schlecht" And Not Selection.Style = "Gut" Then
            ' einen Record vom kassenblatt einlesen
            datBankDatum = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            strBuchungsText = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            intBankkonto = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            varEinnahmen = ActiveCell()
            ActiveCell.Offset(0, 1).Range("A1").Select
            varAusgaben = ActiveCell()
'            varAusgaben = varAusgaben
            
            result = JornalKontoExist(intBankkonto)
            If result = False Then
                MsgBox ("Konto " + CStr(intBankkonto) + " nicht gefunden!")
                Exit Sub
            End If
            C = result  ' column with konto
            Sheets("Jornal").Select
            ' nächste freie zeile ist in Range("JLastEntry") hinterlegt
            intJRow = Range("JPLastRow").Value + 1
            Cells(intJRow, 1).Value = datBankDatum
            Cells(intJRow, 1).NumberFormat = "DD.MM."
            Cells(intJRow, 1).HorizontalAlignment = xlCenter
            Cells(intJRow, 2).Value = strKShort
            Cells(intJRow, 3).Value = strBuchungsText
            
            If varEinnahmen > 0 Then Cells(intJRow, C).Value = varEinnahmen
            If varAusgaben > 0 Then Cells(intJRow, C).Value = varAusgaben
            
            Cells(intJRow + 1, 1).Select
            Selection.EntireRow.Insert
            ' rücksprung nach Tabelle "Ausgangs-Tabelle"
            Sheets(strActiveSheet).Select
            ActiveCell.Offset(0, -4).Range("A1").Select
            ' bearbeitete zeile farblich markieren
            Selection.Style = "Gut"
            ActiveCell.Offset(1, 0).Range("A1").Select
        End If
    Next

EA_Kasse2Jornal_Exit:
    Application.ScreenUpdating = True
    Log ll & "[EOF]"
    Exit Sub

EA_Kasse2Jornal_Error:
    Log ll & Err.Description
    Resume EA_Kasse2Jornal_Exit

End Sub

Sub EA_ToAusEin()
Dim Kommevon, KVon, BankDatum, Buchungstext, BankEin, Betrag, Bankkonto$, Netto, Zeile

'    If gbInitError Then Exit Sub

    Kommevon = Application.ActiveSheet.Name
    'KVon = Left(Application.ActiveSheet.Name, 1)
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select

    BankDatum = ActiveCell()
    ActiveCell.Offset(0, 1).Range("A1").Select
    Buchungstext = ActiveCell()
    ActiveCell.Offset(0, 1).Range("A1").Select
    Bankkonto = ActiveCell()
    ActiveCell.Offset(0, 1).Range("A1").Select
    BankEin = ActiveCell()
    ActiveCell.Offset(0, 1).Range("A1").Select
    Betrag = ActiveCell()
         
    If BankEin > 0 Then
        Betrag = BankEin
        Netto = Betrag / 116 * 100
        
        GoTo ToEin
        Else
        GoTo ToAus
    End If
ToAus:
    Sheets("Ausgaben").Select
    ' nächste freie zeile ist in a1 hinterlegt
    Zeile = Cells(1, 1).Value
    
    Cells(Zeile, 5).Value = BankDatum
    'Cells(nrzeile, 5).NumberFormat = "DD.MM."
    'Cells(nrzeile, 5).HorizontalAlignment = xlCenter
    Cells(Zeile, 2).Value = Buchungstext
    Cells(Zeile, 3).Value = Bankkonto
    Cells(Zeile, 4).Value = Betrag
    
    Cells(1, 1).Value = Zeile + 1
    Cells(Zeile + 1, 1).Select
    GoTo ende
ToEin:
    Sheets("Einnahmen").Select
    ' nächste freie zeile ist in a1 hinterlegt
    Zeile = Cells(1, 1).Value
    
    Cells(Zeile, 3).Value = Buchungstext
    Cells(Zeile, 4).Value = Bankkonto
    Cells(Zeile, 5).Value = Netto
    Cells(Zeile, 8).Value = BankDatum
    'Cells(nrzeile, 5).NumberFormat = "DD.MM."
    'Cells(nrzeile, 5).HorizontalAlignment = xlCenter

    Cells(1, 1).Value = Zeile + 1
    Cells(Zeile + 1, 1).Select
    'ActiveCell.Offset(1, 0).Range("A1").Select
    'Selection.EntireRow.Insert

    ' rücksprung nach Tabelle "komme von"
    'Sheets(KommeVon).Select
    'ActiveCell.Offset(0, -4).Range("A1").Select
    ' bearbeitete zeile farblich markieren
    'With Selection.Interior
    '    .ColorIndex = 40
    '    .Pattern = xlSolid
    'End With

    'ActiveCell.Offset(1, 0).Range("A1").Select
    GoTo ende

ende:
End Sub

Sub EA_KontoSort()
Dim ll$

    ll = lg & "EA_KontoSort > "
    Log ll
    
    If Application.ActiveSheet.Name <> "Konten" Then
        MsgBox "Please select sheet 'Konten' and start again"
        Exit Sub
    
    End If
    
    ActiveSheet.UsedRange.Select
    Range("C1").Activate
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("A2").Select
    
End Sub

Sub EA_nachSA()
Dim intRow%

'    If gbInitError Then Exit Sub
    
    intRow = ActiveCell.row
    Range(Cells(intRow, 1), Cells(intRow, 5)).Select
    Selection.Copy
    Sheets("SA").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Raiba").Select
    Cells(intRow, 2).Select
End Sub

Sub EA_BankFixieren()
Dim ll$, strKShort$
Dim LastRow, FirstRow
Dim rng As Range
Dim Element As Range

    On Error GoTo EA_BankFixieren_Error
    
    ll = lg & "EA_BankFixieren > "
    Log ll
    
    strKShort = GetKontoShort(Application.ActiveSheet.Name)
    If Len(strKShort) = 0 Then
        Err.Raise 900, "EA_BankFixieren", "GetKontoShort failed!"
    End If
    
'    Application.ScreenUpdating = False

    LastRow = Range("B" & strKShort & "LastRow").Value2
    FirstRow = Range("A3").Value
   
    Range("C" & FirstRow & ":F" & LastRow).Select
    Selection.Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Interior.Pattern = xlNone
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Standard"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders.ColorIndex = 38
    
    Range("F" & LastRow).Select
    Selection.Style = "Gut"
    Application.ScreenUpdating = True

EA_BankFixieren_Exit:
    Exit Sub
    
EA_BankFixieren_Error:
    Log ll & Err.Description
    Application.ScreenUpdating = True
    Resume  ' EA_BankFixieren_Exit
    
End Sub

'Function EAP_mw2(Z, Y)
'Dim A, U_B
'
''    If gbInitError Then Exit Function
'
'    A = LTrim(Str(Z))
'    U_B = Left$(A, 2)
'    If U_B = "90" Then
'        EAP_mw2 = Y / 119 * 19
'    End If
'    If U_B = "80" Then
'        EAP_mw2 = Y / 107 * 7
'    End If
'End Function
'
'Function EAP_VorMonat(mon As String)
'Dim Wert, VorMonat
'
''    If gbInitError Then Exit Function
'
'    If mon = 0 Then
'        Wert = 0
'    Else
'        Wert = Worksheets(mon).Range("E23").Value
'    End If
'    VorMonat = Wert
'End Function

Sub EA_ToPurchase()
Dim datDatum As Date
Dim intCol%, n%, intLine As Integer
Dim curBetrag As Currency
Dim ll$, strActiveWb$, strBezeichnung As String
Dim strMonth As String
Dim rng As Range
Dim Element
    
    On Error GoTo EA_ToPurchase_Error
    
    ll = lg & "EA_ToPurchase > "
    Log ll
    
    strActiveWb = ActiveSheet.Name
    
    Set rng = Selection
    intLine = rng.row

    If IsKonto(ActiveSheet.Name) Then
        datDatum = Range("A" & CStr(intLine))
        strMonth = Format(datDatum, "MMMM")       ' Month(datDatum)
        curBetrag = Range("E" & CStr(intLine))
        strBezeichnung = Range("G" & CStr(intLine))
    End If
    
    If ActiveSheet.Name = "Kasse" Then
        datDatum = Range("D" & CStr(intLine))
        strMonth = Format(datDatum, "MMMM")       ' Month(datDatum)
        curBetrag = Range("A" & CStr(intLine))
        strBezeichnung = Range("F" & CStr(intLine))
    End If
    
    Application.ScreenUpdating = False
    
    Sheets("Anschaffungen").Select
    Application.GoTo Reference:="AN" & strMonth
    
    Set rng = Selection
    intLine = rng.row
    intCol = rng.Column
    
    Set rng = Range(Cells(2, intCol + 1), Cells(19, intCol + 1))
    n = 2
    For Each Element In rng
        If Element = "" Then Exit For
        n = n + 1
    Next
    
    Cells(n, intCol) = strBezeichnung
    Cells(n, intCol + 1) = curBetrag

    Sheets(strActiveWb).Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    
EA_ToPurchase_Exit:
    Application.ScreenUpdating = True
    Exit Sub
    
EA_ToPurchase_Error:
    Log ll & Err.Description
    Resume EA_ToPurchase_Exit

End Sub

Sub EA_KasseFixieren()
Dim LastRow, FirstRow
Dim ll As String

    On Error GoTo EA_KasseFixieren_Error
    
    ll = lg & "EA_KasseFixieren > "
    Log ll
    
    If Application.ActiveSheet.Name <> "Kasse" Then
        Err.Raise 900, "EA_KasseFixieren", "Sheet is not a cash account! Exit programm."
    End If
    
    Application.ScreenUpdating = False

    FirstRow = Range("KPFirstRow").row
    LastRow = Range("KPLastRow").Value2

    ' nach Datum sortieren
    Rows(FirstRow & ":" & LastRow).Select
    Selection.Sort Key1:=Range("A" & FirstRow), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ' formeln durch festwerte ersetzen
    Range("C" & FirstRow & ":F" & LastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
        
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders.ColorIndex = 38
    
    Application.CutCopyMode = False
        
    Range("F" & LastRow).Select
    Selection.Interior.ColorIndex = 35
        
    Range("A" & LastRow).Select
    Application.ScreenUpdating = True

EA_KasseFixieren_Exit:
    Exit Sub
    
EA_KasseFixieren_Error:
    MsgBox Err.Description
    Log ll & Err.Description
    Resume EA_KasseFixieren_Exit

End Sub

Function SortStringArray(ByRef strArray As Variant) As Variant()
    'sortieren von String Array
    'eindimensionale Array
    'rekursiver Aufruf
   Dim tmpArray()  As Variant
   Dim i           As Long

    tmpArray = strArray
    
    For i = LBound(tmpArray) To UBound(tmpArray) - 1
        If LCase(strArray(i)) > LCase(strArray(i + 1)) Then
            tmpArray(i) = strArray(i + 1)
            tmpArray(i + 1) = strArray(i)
            strArray = tmpArray
             'rekursiver Aufruf
            tmpArray = SortStringArray(strArray)
        End If
    Next i
     
    SortStringArray = tmpArray
  
End Function

Function SortIntArray(ByRef intArray As Variant) As Variant()
    'sortieren von String Array
    'eindimensionale Array
    'rekursiver Aufruf
   Dim tmpArray()  As Variant
   Dim i           As Long

    tmpArray = intArray
    
    For i = LBound(tmpArray) To UBound(tmpArray) - 1
        If LCase(intArray(i)) > LCase(intArray(i + 1)) Then
            tmpArray(i) = intArray(i + 1)
            tmpArray(i + 1) = intArray(i)
            intArray = tmpArray
             'rekursiver Aufruf
            tmpArray = SortIntArray(intArray)
        End If
    Next i
     
    SortIntArray = tmpArray
  
End Function

Function SelectKonto(konto As Variant)
Dim rng As Range
Dim col

    For Each rng In Range("JPKONTEN")
        If rng.Value = konto Then
            col = rng.Column
            Exit For
        End If
    Next
    
    Cells(2, col).Select
    
End Function

Function GetInsertBeforeAccount(aKonto As Variant) As Variant
Dim ArrayList As Object
Dim konto As Variant
    
    Set ArrayList = CreateObject("System.collections.arraylist")
    
    For Each konto In Range("JPKONTEN")
        ArrayList.Add CInt(konto)
    Next
    
    ArrayList.Sort
    
    Dim bKonto As Integer
    For Each konto In ArrayList
        If konto > aKonto Then
            bKonto = konto
            Exit For
        End If
    Next
    
    GetInsertBeforeAccount = bKonto
    
End Function

Function MainAccountExist(konto As Integer) As Variant
Dim aKonto, mainaccount
Dim rng As Range, rngKonten As Range
Dim col As Integer
Dim blnFound As Boolean

    Set rngKonten = Sheets("Jornal").Range("JPKonten")
    
'    ' search for passed konto
'    aKonto = konto
'    For Each rngK In rng
'        If rngK.Value = aKonto Then
'            col = rngK.Column
'            blnFound = True
'            Exit For
'        End If
'    Next
'    ' Hauptkonto not found, try subkonto
'    If blnFound = False Then
'        aKonto = Int(konto / 10) * 10
'        For Each rngK In rng
'            If rngK.Value = aKonto Then
'                col = rngK.Column
'                blnFound = True
'                Exit For
'            End If
'        Next
'    End If
'
'
    ' Hauptkonto not found, try main konto
    If blnFound = False Then
        mainaccount = Int(konto / 100) * 100
        For Each rng In rngKonten
            aKonto = CInt(rng.Value)
            col = rng.Column
            If aKonto = mainaccount Then
                blnFound = True
                Exit For
            End If
        Next
    End If
    
    Dim ar(3) As Variant
    
    ar(0) = blnFound
    ar(1) = mainaccount
    ar(2) = col
    
    MainAccountExist = ar
    
    
End Function

Sub test()
Dim result

    result = MainAccountExist(1420)
    
    Debug.Print result(0), result(1), result(2)
    
End Sub

Function JornalKontoExist(intKonto As Integer) As Variant
Dim col%, row As Integer
Dim comefrom$, bezeichnung As String
Dim blnFound As Boolean
Dim result, mainaccount
Dim rng As Range
Dim kontoafter, konto, mkonto
    
    On Error GoTo JornalKontoExist_Error
    
    JornalKontoExist = False
    blnFound = False
    
    For Each rng In Sheets("Jornal").Range("JPKONTEN")
        If rng.Value = intKonto Then
            blnFound = True
            JornalKontoExist = rng.Column
            Exit Function
        End If
    Next
    
    mainaccount = MainAccountExist(intKonto)
    
    blnFound = mainaccount(0)
    mkonto = mainaccount(1)
    col = mainaccount(2)
    
    If blnFound = False Then
        result = MsgBox("Konto " & intKonto & " nicht gefunden!" & vbLf & "Wollen Sie das Konto neu erstellen?", vbYesNo, "JornalKontoExist")
        If result = vbYes Then ' neues konto erstellen
            bezeichnung = InputBox("Bitte Kontobezeichnung eingeben", "Neues Konto", "")
            
            
            comefrom = ActiveSheet.Name
            Sheets("Jornal").Select
            kontoafter = GetInsertBeforeAccount(intKonto)
            SelectKonto kontoafter
            
            Application.DisplayAlerts = False
            
            row = ActiveCell.row
            col = ActiveCell.Column
            
            Selection.EntireColumn.Insert
            Cells(1, col).Select
            With Selection
                .Value2 = bezeichnung
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Cells(2, col).Select
            With Selection
                .Value2 = intKonto
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
                
            JornalKontoExist = col
        Else
            JornalKontoExist = False
        End If
    Else
        JornalKontoExist = col
    End If

JornalKontoExist_Exit:
    If comefrom <> "" Then Sheets(comefrom).Activate
    Exit Function

JornalKontoExist_Error:
    JornalKontoExist = False
    Debug.Print Err.Description
    Resume JornalKontoExist_Exit
    
End Function

Function ExpandNames()
Dim wb As Workbook
Dim ws As Worksheet
Dim nameString As String
Dim token As String
Dim xName As Name
Dim nameIndex As Integer
Dim ar

    Set wb = Application.ActiveWorkbook
    Set ws = wb.Worksheets("Jornal")
    
    nameString = "Jornal!JPKonten"
    
    For Each xName In ws.Names
'        Debug.Print xName.Name, xName.Index, xName.RefersTo
        If xName.Name = nameString Then
            nameIndex = xName.Index
            Exit For
        End If
    Next
    
    ar = Split(xName.RefersTo, ":")
    token = ar(1)
    Set xName = wb.Names.Item(nameIndex)

    xName.RefersTo = "=Jornal!$D$2:" & token
End Function

Function GetKontoShort(strKonto As String) As String
    If Left(strKonto, 1) = "K" Then
        GetKontoShort = Mid(strKonto, 2, 1)
    End If
End Function

Function GetKontoName(strKonto As String) As String
    
    If Left(strKonto, 1) = "K" Then
        GetKontoName = Mid(strKonto, 4, 10)
    End If
    
End Function

Function IsKonto(strKonto As String) As Boolean
    
    If Left(strKonto, 1) = "K" Then
        IsKonto = True
    Else
        IsKonto = False
    End If
    
End Function

Sub EA_Test()
    Debug.Print GetFirstOfMonth(3)
    Sheets("Jornal").Selection.EntireRow.Insert
    
End Sub

