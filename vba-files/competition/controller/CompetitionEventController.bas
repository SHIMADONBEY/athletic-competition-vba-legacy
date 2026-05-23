Attribute VB_Name = "CompetitionEventController"
'namespace=vba-files/competition/controller
Option Explicit
Option Private Module

Private Enum CompetitionEventsColumnEnum
    IDX_ID = 1
    IDX_CATEGORY
    IDX_SEX
    IDX_EVENT_NAME
    IDX_SUPECIFICATION
    IDX_EVENT_TYPE
    IDX_ROUNDS_COUNT
    IDX_PERSONS_PER_GROUP
    IDX_OPERATION_WIND_GAUGE
    IDX_MEASUREMENT_RT
    IDX_MINUTE_PER_GROUP
    IDX_CALL_START_TIMING
    IDX_CALL_END_TIMING
    IDX_TEMPLATE_FILE_NAME = 16
    IDX_OUTPUT_FILE_NAME
End Enum

Private Const HEADER_ROW As Long = 3

Public Sub ExportCompetitionEvents()
    Dim vCompetionSettings As CompetitionConfigModel: Set vCompetionSettings = CompetitionConfigFactory.GetConfigInstance()
    Dim vCompetionEvents As CompetitionEventModels: Set vCompetionEvents = CompetitionEventRepository.ReadAllEvents(true)

    Dim vCompetionEvent As CompetitionEventModel
    Dim vAllCount As Variant: vAllCount = CDec(vCompetionEvents.Count())
    Dim vProcessCount As Long: vProcessCount = 0
    For Each vCompetionEvent In vCompetionEvents.All()
        Dim vProgress As Variant: vProgress = Fix((CDec(vProcessCount) * CDec(1000) + CDec(5)) / (vAllCount * CDec(10)))
        Application.StatusBar = MessageFactory.Generate("SI007").Prompt(vCompetionEvent.EventKey(), vProgress)

        Call CompetitionEventExtractor.ExtractCompetionEventToFile(vCompetionSettings, vCompetionEvent)

        vProcessCount = vProcessCount + 1
    Next vCompetionEvent

    Application.StatusBar = False
End Sub

Public Sub UpdateCompetitionEvents()
    Call CompetitionEventRepository.ReadAllEvents(true)
End Sub

Public Sub OnEventNameUpdate(Spot As Range)
    If (Application.Intersect(Spot, Range("CompetitionEventList").Columns(CompetitionEventsColumnEnum.IDX_EVENT_NAME)) Is Nothing) Then
        Exit Sub
    End If

    If (CodeMasterRepository.EventMaster.FindByName(Spot.Value) Is Nothing) Then
        Exit Sub
    End If

    Dim tableRowIndex As Long: tableRowIndex = Spot.Row - HEADER_ROW
    Dim vEventMasterData As EventMasterModel: Set vEventMasterData = CodeMasterRepository.EventMaster.FindByName(Spot.Value)
    Range("CompetitionEventList").Cells(tableRowIndex, CompetitionEventsColumnEnum.IDX_EVENT_TYPE).Value = vEventMasterData.Category
    Range("CompetitionEventList").Cells(tableRowIndex, CompetitionEventsColumnEnum.IDX_OPERATION_WIND_GAUGE).Value = CodeMasterRepository.WindScalingCode.Item(vEventMasterData.NeedWindMeasuring)
End Sub

Private  Function ThisSheet() As Worksheet
    Set ThisSheet = Sheets("Ží–Ú")
End Function
