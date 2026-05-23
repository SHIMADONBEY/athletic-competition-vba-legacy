VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "RoundListSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub BtnRoundExportClicked()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_btn_round_export_clicked
        Call CompetitionRoundController.ExportCompetitionRounds()
        Call MessageFactory.Generate("SI008").ToMsgBox()
    GoTo finally_btn_round_export_clicked
    on_failure_btn_round_export_clicked:
        Select Case Err.Number
            Case CustomErrorCodeEnum.EventFileNotFound
                MsgBox Err.Description, vbCritical + vbOKOnly, MessageFactory.Generate("SE009").Title
            Case Else 
                Call MessageFactory.Generate("SF001").ToMsgBox()
        End Select
    finally_btn_round_export_clicked:
        Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Deactivate()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failed_sheet_deactivate
        Call CompetitionRoundController.UpdateRoundInfo()
    GoTo finally_sheet_deactivate
on_failed_sheet_deactivate:
    Me.Activate
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Select Case Err.Number 
        Case CustomErrorCodeEnum.DuplicatedRound 
            MsgBox Err.Description, vbCritical + vbOKOnly, MessageFactory.Generate("SE002").Title
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select

    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_sheet_deactivate:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failed_cell_change

        Dim spotRange As Range
        For Each spotRange In Target
            Call CompetitionRoundController.OnUpdateStartTime(spotRange)
            Call CompetitionRoundController.OnUpdateGroupCount(spotRange)
        Next spotRange 

    GoTo fainally_cell_change
on_failed_cell_change:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
fainally_cell_change:
    Call MacroProcessorFactory.DestroyInstance()
End Sub
