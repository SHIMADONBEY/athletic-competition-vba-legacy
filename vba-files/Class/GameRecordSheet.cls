VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "GameRecordSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Worksheet_Deactivate()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_worksheet_deactivate
        Call GameRecordController.UpdateGameRecords()
    GoTo finally_worksheet_deactivate
    on_failure_worksheet_deactivate:
        Me.Activate
        Call MessageFactory.Generate("SF001").ToMsgBox()
    finally_worksheet_deactivate:
        Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnGameRecordsExport_Clicked()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_btn_game_records_export_clicked
        Call GameRecordController.ExportGameRecords()
        Call MessageFactory.Generate("SI014").ToMsgBox()
    GoTo finally_btn_game_records_export_clicked
    on_failure_btn_game_records_export_clicked:
        Me.Activate
        Select Case Err.Number
            Case CustomErrorCodeEnum.EventFileNotFound
                MsgBox Err.Description, vbCritical + vbOKOnly, MessageFactory.Generate("SE009").Title
            Case Else 
                Call MessageFactory.Generate("SF001").ToMsgBox()
        End Select

    finally_btn_game_records_export_clicked:
        Call MacroProcessorFactory.DestroyInstance()
End Sub
