Attribute VB_Name = "modMenu"
Option Explicit

Const lg = "modMenu > "


Sub AUTools_MenuDelete()
    On Error Resume Next
    With Application.CommandBars(1)
        .Controls("AUTools").Delete
    End With
End Sub

Sub AUTools_MenuInsert()
Dim cbMenu As CommandBar
Dim AUToolsMenu As CommandBarControl
Dim cbCommand As CommandBarControl
Dim i%, anz As Integer
Dim ll$

    ll = lg & "AUTools_MenuInsert > "
    Log ll
    
    anz = Application.CommandBars(1).Controls.Count
    For i = 1 To anz - 1
        If Application.CommandBars(1).Controls(i).Caption = "AUTools" Then
            Exit Sub
        End If
    Next i
  
    Set cbMenu = CommandBars.ActiveMenuBar
    ' set mainmenu
    Set AUToolsMenu = cbMenu.Controls.Add(Type:=msoControlPopup, Temporary:=False, Before:=anz)
        AUToolsMenu.Caption = "AUTools"
        AUToolsMenu.TooltipText = "click here to open the AUTools"
    '
    Log ll & "insert menuitem new AUTools"
    Set cbCommand = AUToolsMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Tools Info"
        .Style = msoButtonCaption
        .OnAction = "GUIShow_AUToolsInfo"
        .Enabled = True
        .Visible = True
    End With
    '
    Log ll & "insert menuitem new Test"
    Set cbCommand = AUToolsMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "Test only"
        .Style = msoButtonCaption
        .OnAction = "MarkCostOrg"
        .Enabled = False
        .Visible = False
    End With
    '
    Log ll & "insert menuitem About"
    Set cbCommand = AUToolsMenu.Controls.Add(Type:=msoControlButton)
    With cbCommand
        .Caption = "About"
        .Style = msoButtonCaption
        .OnAction = "GUIShow_AUToolsAbout"
        .BeginGroup = True
        .Enabled = True
        .Visible = True
        .TooltipText = ThisWorkbook.FullName
    End With
End Sub

Sub AUTools_MenuActivate()
Dim i%, anz As Integer

    anz = Application.CommandBars(1).Controls.Count
    For i = 1 To anz - 1
        If Application.CommandBars(1).Controls(i).Caption = "AUTools" Then
            Application.CommandBars(1).Controls(i).Enabled = True
            Exit Sub
        End If
    Next i
End Sub

Sub AUTools_MenuDeactivate()
Dim i%, anz As Integer

    anz = Application.CommandBars(1).Controls.Count
    For i = 1 To anz - 1
        If Application.CommandBars(1).Controls(i).Caption = "AUTools" Then
            Application.CommandBars(1).Controls(i).Enabled = False
            Exit Sub
        End If
    Next i
End Sub

