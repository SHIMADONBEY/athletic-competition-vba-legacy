Attribute VB_Name = "GameRecordRepository"
'namespace=vba-files/competition/repository
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

Private m_Repository As GameRecordModels

Public Function ReadAllGameRecords(Optional Reload As Boolean = False) As GameRecordModels
    If (Reload Or m_Repository Is Nothing) Then
        Dim vRecords As GameRecordModels: Set vRecords = New GameRecordModels
        Dim vRowRange As Range
        For Each vRowRange In Range("GameRecordList").Rows()
            Dim vRecord As GameRecordModel: Set vRecord = ReadFromRow(vRowRange)
            If (vRecord Is Nothing) Then 
                ' DO NOTHING
            Else
                Call vRecords.Add(vRecord)
                vRowRange.Cells(1, GameRecordColumnEnum.IDX_ID).Value = vRecord.Id
            End If
        Next vRowRange

        Set m_Repository = vRecords
    End If

    Set ReadAllGameRecords = m_Repository
End Function

Private Function ReadFromRow(RowRange As Range) As GameRecordModel
    Set ReadFromRow = Nothing

    Dim pEventKey As String     : pEventKey = RowRange.Cells(1, GameRecordColumnEnum.IDX_EVENT).Value
    Dim pName As String         : pName = RowRange.Cells(1, GameRecordColumnEnum.IDX_NAME).Value
    Dim pRecordString As String : pRecordString = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORD).Value
    Dim pAthleteName As String  : pAthleteName = RowRange.Cells(1, GameRecordColumnEnum.IDX_ATHLETE_NAME).Value
    Dim pTeamName As String     : pTeamName = RowRange.Cells(1, GameRecordColumnEnum.IDX_TEAM_NAME).Value

    If (pEventKey = "" Or pName = "" Or pRecordString = "" Or pAthleteName = "" Or pTeamName = "") Then
        Exit Function
    End If

    Dim vEvent As CompetitionEventModel: Set vEvent = CompetitionEventRepository.ReadAllEvents().Item(RowRange.Cells(1, GameRecordColumnEnum.IDX_EVENT_ID).Value)
    If (vEvent Is Nothing) Then
        Exit Function
    End If

    Dim vRecordValue As IResultValue: Set vRecordValue = ResultValueParser.Parse(pRecordString, vEvent.EventType, CompetitionConfigFactory.GetConfigInstance().MinuteDelimiter, CompetitionConfigFactory.GetConfigInstance().SecondDelimiter, CompetitionConfigFactory.GetConfigInstance().MeterDelimiter)
    If (vRecordValue Is Nothing) Then
        Exit Function
    ElseIf (vRecordValue.Valid = False) Then
        Exit Function
    End If

    Dim vRecordedYmd As String
    Dim pRecordedYear As String     : pRecordedYear = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORDED_YEAR).Value
    Dim pRecordedMonth As String    : pRecordedMonth = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORDED_MONTH).Value
    Dim pRecordedDay As String      : pRecordedDay = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORDED_DAY).Value

    If (pRecordedYear = "") Then 
        vRecordedYmd = ""
    ElseIf (pRecordedMonth = "" Or pRecordedDay = "") Then
        vRecordedYmd = pRecordedYear
    Else
        vRecordedYmd = pRecordedYear & "/" & pRecordedMonth & "/" & pRecordedDay
    End If

    With New GameRecordModel
        Call .Initialize( _
                vEvent.Id _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_SERIAL).Value _
                , pName _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_ABBREBIATION).Value _
                , vRecordValue _
                , pAthleteName _
                , pTeamName _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TEAM_PLACE).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TEAM_COUNTRY).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORDED_YEAR).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORDED_MONTH).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORDED_DAY).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_SCALE).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TARGET_GRADE).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_COMMENT).Value _
                , vRecordedYmd _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_ID).Value _        
        )

        Set ReadFromRow = .Self()
    End With
End Function
