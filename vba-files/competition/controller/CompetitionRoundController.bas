Attribute VB_Name = "CompetitionRoundController"
'namespace=vba-files/competition/controller
Option Explicit
Option Private Module

Private Enum CompetitionRoundColumnEnum
    IDX_EVENT_ID = 1
    IDX_EVENT_NAME
    IDX_ROUND_ID
    IDX_NAME
    IDX_START_DATE_TIME
    IDX_ORDER_COUNT
    IDX_GROUP_COUNT
    IDX_GROUP_STRATEGY
    IDX_ORDER_STRATEGY
    IDX_PLACES_PER_GROUP
    IDX_QUALIFYING_STANDARD
    IDX_ADDITION_COUNT
    IDX_EVENT_AREA
    IDX_END_DATE_TIME
    IDX_GROUPS_PER_CALL
    IDX_GROUP_STRATEGY_CODE
    IDX_ORDER_STRATEGY_CODE
    IDX_ID
End Enum

Private Const HEADER_ROW As Long = 3

Private Function ThisSheet() As Worksheet
    Set ThisSheet = Sheets("ƒ‰ƒEƒ“ƒh")
End Function

Public Sub ExportCompetitionRounds()
    Dim vCompetitionEvents As CompetitionEventModels: Set vCompetitionEvents = CompetitionEventRepository.ReadAllEvents()
    Dim vEventRounds As CompetitionRoundModels: Set vEventRounds = CompetitionRoundRepository.ReadAllRounds(True)

    Dim vRoundsEventGroup As Object: Set vRoundsEventGroup = vEventRounds.GroupByEvents()
    Dim vAllCount As Variant: vAllCount = CDec(vEventRounds.Count())
    Dim vProcessCount As Long: vProcessCount = 0

    Dim vEvent As CompetitionEventModel
    For Each vEvent In vCompetitionEvents.All()
        If (vRoundsEventGroup.Exists(vEvent.Id)) Then
            Dim vProgress As Variant: vProgress = Fix((CDec(vProcessCount * 1000) / vAllCount + CDec(5)) / CDec(10))
            Application.StatusBar = MessageFactory.Generate("SI010").Prompt(vEvent.EventKey(), vProgress)

            Dim vRounds As CompetitionRoundModels: Set vRounds = vRoundsEventGroup.Item(vEvent.Id)
            Call CompetitionRoundLoader.LoadCompetitionRoundToFile(vEvent, vRounds)

            vProcessCount = vProcessCount + vRounds.Count()
        End If
    Next vEvent

    Application.StatusBar = False
End Sub

Public Sub UpdateRoundInfo()
    Call CompetitionRoundRepository.ReadAllRounds(true)
End Sub

Public Sub OnUpdateGroupCount(Spot As Range)
    If (Application.Intersect(Spot, Range("CompetitionRoundList").Columns(CompetitionRoundColumnEnum.IDX_GROUP_COUNT)) Is Nothing) Then
        Exit Sub
    End If

    Dim tableRowIndex As Long: tableRowIndex = Spot.Row - HEADER_ROW
    Call UpdateEndTime(Range("CompetitionRoundList").Rows(tableRowIndex))
End Sub

Public Sub OnUpdateStartTime(Spot As Range)
    If (Application.Intersect(Spot, Range("CompetitionRoundList").Columns(CompetitionRoundColumnEnum.IDX_START_DATE_TIME)) Is Nothing) Then
        Exit Sub
    End If

    Dim tableRowIndex As Long: tableRowIndex = Spot.Row - HEADER_ROW
    Call UpdateEndTime(Range("CompetitionRoundList").Rows(tableRowIndex))
End Sub

Private Sub UpdateEndTime(rowRange As Range)
    Dim eventId As Long
    If (IsNumeric(rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_EVENT_ID).Value)) Then 
        eventId = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_EVENT_ID).Value
    Else
        rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_END_DATE_TIME).Value = Empty
        Exit Sub
    End If
    
    Dim eventData As CompetitionEventModel: Set eventData = CompetitionEventRepository.ReadAllEvents().Item(eventId)
    If (eventData Is Nothing) Then
        rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_END_DATE_TIME).Value = Empty
        Exit Sub
    End If

    Dim mStartTime As Date: mStartTime = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_START_DATE_TIME)
    Dim mGroupCount As Long: mGroupCount = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_GROUP_COUNT)

    rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_END_DATE_TIME).Value = mStartTime + mGroupCount * eventData.MinutesPerGroup / 1440
End Sub
