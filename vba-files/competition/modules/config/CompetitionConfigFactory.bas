Attribute VB_Name = "CompetitionConfigFactory"
'namespace=vba-files/competition/modules/config
Option Explicit
Option Private Module

Private Enum CompetionSettingRowEnum
    IDX_COMPETITION_NAME = 4
    IDX_COMPETITION_LATIN
    IDX_COMPETITION_CODE
    IDX_FACILITY_NAME = 8
    IDX_FACILITY_LATIN
    IDX_FACILITY_CODE
    IDX_FACILITY_PLACE
    IDX_COMPETITION_YEAR = 13
    IDX_COMPETITION_DATE_START
    IDX_COMPETITION_DATE_END
    IDX_COMPETIBLE_EVENTS = 38
    IDX_LANES_COUNT
    IDX_MINUTE_DELIMITER = 41
    IDX_SECOND_DELIMITER
    IDX_METER_DELIMITER
End Enum

Private Const SPONSORS_LIST_TOP_ROW As Long = 17
Private Const SPONSORS_LIST_END_ROW As Long = 23

Private m_ConfigInstance As CompetitionConfigModel

Private Function ThisSheet() As Worksheet
    Set ThisSheet = Sheets("ëÂâÔê›íË")
End Function

Public Sub UpdateConfig()
    ' äÓñ{èÓïÒÇÃê›íË
    Dim basicInfoData As CompetitionInfoModel: Set basicInfoData = New CompetitionInfoModel

    Call basicInfoData.Initialize( _
            ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETITION_NAME, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETITION_LATIN, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETITION_CODE, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_FACILITY_NAME, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_FACILITY_LATIN, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_FACILITY_CODE, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_FACILITY_PLACE, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETITION_YEAR, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETITION_DATE_START, 6).value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETITION_DATE_END, 6).value _
    )
    ' Debug.Print basicInfoData.ToJson()

    ' éÂç√ícëÃÇÃê›íË
    Dim sponsorList As SponsorModels: Set sponsorList = New SponsorModels
    Dim readRow As Long
    For readRow = SPONSORS_LIST_TOP_ROW To SPONSORS_LIST_END_ROW
        With New SponsorModel
            Call .Initialize(ThisSheet.Cells(readRow, 4).Value, ThisSheet.Cells(readRow, 6).Value)
            Call sponsorList.Add(.Self())
        End With
    Next readRow

    ' Debug.Print sponsorList.ToJson()

    ' TODO: ìæì_ê›íË
    Dim scoringConfig As CompetitionScoringModel: Set scoringConfig = New CompetitionScoringModel

    Set m_ConfigInstance = New CompetitionConfigModel
    Call m_ConfigInstance.Initialize( _
            basicInfoData _
            , sponsorList _
            , scoringConfig _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_COMPETIBLE_EVENTS, 6).Value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_LANES_COUNT, 6).Value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_MINUTE_DELIMITER, 6).Value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_SECOND_DELIMITER, 6).Value _
            , ThisSheet.Cells(CompetionSettingRowEnum.IDX_METER_DELIMITER, 6).Value _
    )
End Sub

Public Function GetConfigInstance(Optional Reload As Boolean = True) As CompetitionConfigModel
    If (Reload Or m_ConfigInstance Is Nothing) Then
        Call UpdateConfig()
    End If
    Set GetConfigInstance = m_ConfigInstance
End Function
