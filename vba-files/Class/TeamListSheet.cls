VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TeamListSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_Deactivate()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_update_teams
        Call TeamListController.UpdateTeamList()
    GoTo finally_update_teams
on_failure_update_teams:
    Me.Activate
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_update_teams:
    Call MacroProcessorFactory.DestroyInstance()
End Sub
