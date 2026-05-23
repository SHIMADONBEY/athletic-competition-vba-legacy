Attribute VB_Name = "CompetitionRoundRepository"
'namespace=vba-files/competition/repository
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

Private m_Records As CompetitionRoundModels

Private Function ThisSheet() As Worksheet
    Set ThisSheet = Sheets("ラウンド")
End Function

Public Function ReadAllRounds(Optional Reload As Boolean = False) As CompetitionRoundModels
    If (Reload Or m_Records Is Nothing) Then
        Set m_Records = New CompetitionRoundModels
        Dim rowRange As Range
        For Each rowRange In Range("CompetitionRoundList").Rows()
            Dim mEventName As String: mEventName = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_EVENT_NAME).Value
            Dim mRoundId As Long: mRoundId = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_ROUND_ID).Value
            Dim mRoundName As String: mRoundName = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_NAME).Value

            If Not (mEventName = "" Or IsEmpty(mRoundId) Or mRoundName = "") Then
                Dim mEventId As Long: mEventId = rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_EVENT_ID).Value
                If Not (m_Records.FindByEventRound(mEventId, mRoundId) Is Nothing) Then
                    ' ラウンド重複のためエラー
                    Err.Raise CustomErrorCodeEnum.DuplicatedRound, "CompetitionRoundRepository", MessageFactory.Generate("SE002").Prompt(mEventName, mRoundId)
                End If

                With New CompetitionRoundModel
                    Call .Initialize( _
                            mEventId _
                            , mRoundId _
                            , mRoundName _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_START_DATE_TIME).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_ORDER_COUNT).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_GROUP_COUNT).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_GROUP_STRATEGY_CODE).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_ORDER_STRATEGY_CODE).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_PLACES_PER_GROUP).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_QUALIFYING_STANDARD).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_ADDITION_COUNT).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_EVENT_AREA).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_END_DATE_TIME).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_GROUPS_PER_CALL).Value _
                            , rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_ID).Value _
                    )

                    Call m_Records.Add(.Self())
                    rowRange.Cells(1, CompetitionRoundColumnEnum.IDX_ID).Value = .Id
                End With
            End If
        Next rowRange 
    End If

    ' Debug.Print m_Records.ToJson()
    Set ReadAllRounds = m_Records
End Function