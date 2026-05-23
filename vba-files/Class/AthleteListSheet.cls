VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "AthleteListSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Sub BtnEntriesExportClicked()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_btn_entries_export_clicked
        Call AthletePersonController.ExportEntries()
        Call MessageFactory.Generate("SI011").ToMsgBox()
    GoTo finally_btn_entries_export_clicked
    on_failure_btn_entries_export_clicked:
        Select Case Err.Number
            Case CustomErrorCodeEnum.EventFileNotFound
                MsgBox Err.Description, vbCritical + vbOKOnly, MessageFactory.Generate("SE009").Title
            Case Else 
                Call MessageFactory.Generate("SF001").ToMsgBox()
        End Select
    finally_btn_entries_export_clicked:
        Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Deactivate()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failure_update_entries
        Call AthletePersonController.UpdateAll()
    GoTo finally_update_entries
on_failure_update_entries:
    Me.Activate
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_update_entries:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'     Call MacroProcessorFactory.CreateInstance()
'     On Error GoTo on_failure_value_change
'         Dim vSpotCell As Range
'         For Each vSpotCell In Target
'         Next vSpotCell 
'     GoTo finally_value_change
' on_failure_value_change:
'     Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
'     Call MessageFactory.Generate("SF001").ToMsgBox()
'     Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
' finally_value_change:
'     Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
 