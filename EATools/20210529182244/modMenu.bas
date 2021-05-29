Attribute VB_Name = "modMenu"
Option Explicit

'
' EA modMenu
' ==========
'
Const lg As String = "modMenu > "

Sub EA_MenuDelete()
Dim ll$

    ll = lg & "EA_MenuDelete > "
    
    On Error Resume Next
    
    Log ll
    Log ll & "delete EAP Menu"

    With Application.CommandBars(1)
        .Controls("EATools").Delete
        ' .Controls("E+A-").Delete
        ' .Controls("E+A-" & gcstrKlasse).Delete
        ' .Controls("E+A-&" & gcstrKlasse).Delete
        ' .Controls("Anwender-Tools").Delete
    End With
   
    Log ll & "[EOF]"
    
EA_MenuDelete_Exit:

End Sub

Sub EA_MenuInsert()
Dim cbMenu As CommandBar
Dim EAMenu As CommandBarControl
Dim cbCommand As CommandBarControl
Dim i%, anz As Integer
Dim ll$

    On Error GoTo EA_MenuInsert_Error
    
    ll = lg & "EA_MenuInsert > "
    Log ll

    anz = Application.CommandBars(1).Controls.Count
    For i = 1 To anz - 1
        If Application.CommandBars(1).Controls(i).Caption = gcstrAIShort Then
            Err.Raise 900, "EA_MenuInsert", "menu " & gcstrAIShort & " allready installed."
        End If
    Next i
    ' menu not found, create menu
    Log ll & "createing Menu " & gcstrAIShort
    
    Set cbMenu = CommandBars.ActiveMenuBar
    ' set mainmenu
    Set EAMenu = cbMenu.Controls.Add(Type:=msoControlPopup, Temporary:=False, before:=anz)
        EAMenu.Caption = "EATools"  ' & gcstrKlasse
        EAMenu.TooltipText = ADDINNAME
    ' set ersten Menüpunkt
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Bank fixieren"
        .OnAction = "EA_BankFixieren"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Bank 2 Jornal"
        .Style = msoButtonCaption
        ' .OnAction = "EA" & gcstrKlasse & "_Bank2Jornal"
        .OnAction = "EA_Bank2Jornal"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Kasse fixieren"
        .OnAction = "EA_KasseFixieren"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Kasse 2 Jornal"
        .Style = msoButtonCaption
        ' .OnAction = "EA" & gcstrKlasse & "_Kasse2Jornal"
        .OnAction = "EA_Kasse2Jornal"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Konten sortieren"
        .Style = msoButtonCaption
        .OnAction = "EA_KontoSort"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
        .BeginGroup = True
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "New Template"
        .Style = msoButtonCaption
        .OnAction = "EA_NewTemplate"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
        .BeginGroup = True
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Reset Keys"
        .Style = msoButtonCaption
        .OnAction = "EA_ResetKey"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Menü löschen"
        .Style = msoButtonCaption
        .OnAction = "EA_MenuDelete"
        .Enabled = True
        .Visible = True
        .TooltipText = ""
        .BeginGroup = True
    End With
    Set cbCommand = EAMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "About"
        .Style = msoButtonCaption
        .OnAction = "EA_About"
        .Enabled = True
        .Visible = True
        .TooltipText = ThisWorkbook.FullName
        .BeginGroup = True
    End With
    
EA_MenuInsert_Exit:
    Log ll & "[EOF]"
    Exit Sub

EA_MenuInsert_Error:
    Log ll & "Error " & Err.Number & ": " & Err.Description
    Resume EA_MenuInsert_Exit

End Sub

Sub EA_MenuActivate()
Dim i%, anz As Integer
Dim ll$

    ll = lg & "EA_MenuActivate > "
    Log ll
    Log ll & "activate EA Menu"

    anz = Application.CommandBars(1).Controls.Count
    For i = 1 To anz - 1
        If Application.CommandBars(1).Controls(i).Caption = "EATools" Then
            Application.CommandBars(1).Controls(i).Enabled = True
            Exit For
        End If
    Next i
    Log ll & "[EOF]"

End Sub

Sub EA_MenuDeactivate()
Dim i%, anz As Integer
Dim ll$

    ll = lg & "EA_MenuDeactivate > "
    Log ll
    Log ll & "deactivate EA Menu"

    anz = Application.CommandBars(1).Controls.Count
    For i = 1 To anz - 1
        If Application.CommandBars(1).Controls(i).Caption = "EATools" Then
            Application.CommandBars(1).Controls(i).Enabled = False
            Exit For
        End If
    Next i
    Log ll & "[EOF]"

End Sub

