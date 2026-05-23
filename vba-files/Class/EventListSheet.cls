VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "EventListSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub BtnEventExport_Clicked()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_btn_event_export_clicked
        Call CompetitionEventController.ExportCompetitionEvents()

        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        Call MessageFactory.Generate("SI004").ToMsgBox()
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    GoTo finally_btn_event_export_clicked
on_failure_btn_event_export_clicked:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Select Case Err.Number 
        Case CustomErrorCodeEnum.NotApplicated
            MsgBox Err.Description, vbCritical + vbOKOnly, MessageFactory.Generate("SE005").Title 
        Case CustomErrorCodeEnum.TemplateFileNotFound
            MsgBox Err.Description, vbCritical + vbOKOnly, MessageFactory.Generate("SE006").Title 
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select
    
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_btn_event_export_clicked:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Deactivate()
    Call CompetitionEventController.UpdateCompetitionEvents()
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call MacroProcessorFactory.CreateInstance()
    
    Dim spotCell As Range
    For Each spotCell In Target
        Call CompetitionEventController.OnEventNameUpdate(spotCell)
    Next spotCell 
    GoTo finally_worksheet_change
on_failed_worksheet_change:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_worksheet_change:
    Call MacroProcessorFactory.DestroyInstance()
End Sub