VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AllUserTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const lg As String = "Workbook > "

Private Sub Workbook_Open()

    Call AUTools_SplashTimer(2)
    Call AUTools_MenuDelete
    Call AUTools_MenuInsert
    
End Sub
