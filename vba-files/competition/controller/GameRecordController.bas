Attribute VB_Name = "GameRecordController"
'namespace=vba-files/competition/controller
Option Explicit
Option Private Module

Private Enum GameRecordColumnEnum
    IDX_EVENT = 1
    IDX_SERIAL
    IDX_NAME
    IDX_ABBREBIATION
    IDX_RECORD
    IDX_SCALE_INPUT
    IDX_ATHLETE_NAME
    IDX_TEAM_NAME
    IDX_TEAM_PLACE
    IDX_TEAM_COUNTRY
    IDX_RECORDED_YEAR
    IDX_RECORDED_MONTH
    IDX_RECORDED_DAY
    IDX_TARGET_GRADE
    IDX_COMMENT
    IDX_ID
    IDX_EVENT_ID
    IDX_SCALE
End Enum

Private Const HEADER_ROW As Long = 3

Public Sub UpdateGameRecords()
    Call GameRecordRepository.ReadAllGameRecords(True)
End Sub

Public Sub ExportGameRecords()
    Dim vCompetitionEvents As CompetitionEventModels: Set vCompetitionEvents = CompetitionEventRepository.ReadAllEvents(true)
    Dim vGameRecords As GameRecordModels: Set vGameRecords = GameRecordRepository.ReadAllGameRecords(true)

    Dim vGameRecordsEventGroup As Object: Set vGameRecordsEventGroup = vGameRecords.GroupBy(New GameRecordEventClassifier)
    Dim vAllCount As Variant: vAllCount = CDec(vGameRecords.Count())
    Dim vProgressCount As Long: vProgressCount = 0

    Dim vEvent As CompetitionEventModel
    For Each vEvent In vCompetitionEvents.All()
        If (vGameRecordsEventGroup.Exists(vEvent.Id)) Then
            Dim vProgress As Variant: vProgress = Fix(CDec((vProgressCount * 1000) / vAllCount + CDec(5)) / CDec(10))
            Application.StatusBar = MessageFactory.Generate("SI013").Prompt(vEvent.EventKey(), vProgress)

            Dim vGameRecordList As GameRecordModels: Set vGameRecordList = vGameRecordsEventGroup.Item(vEvent.Id)
            Call GameRecordLoader.LoadGameRecordsToFile(vEvent, vGameRecordList)

            vProgressCount = vProgressCount + vGameRecordList.Count()
        End If
    Next vEvent

    Application.StatusBar = False
End Sub
