Attribute VB_Name = "CompetitionEventRepository"
'namespace=vba-files/competition/repository
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
    IDX_MINUTE_PER_GROUP = 11
    IDX_CALL_START_TIMING
    IDX_CALL_END_TIMING
    IDX_TEMPLATE_FILE_NAME = 16
    IDX_OUTPUT_FILE_NAME
    IDX_EVENT_TYPE_CODE
    IDX_OPERATION_WIND_GAUGE
    IDX_MEASUREMENT_RT
End Enum

Private Const HEADER_ROW As Long = 3

Private m_Records As CompetitionEventModels

Public Function ReadAllEvents(Optional Reload As Boolean = False) As CompetitionEventModels
    If (Reload Or m_Records Is Nothing) Then
        Set m_Records = New CompetitionEventModels

        Dim rowRange As Range
        For Each rowRange In Range("CompetitionEventList").Rows
            If Not (rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_ID) = "") Then
                With New CompetitionEventModel
                    Call .Initialize( _
                            rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_ID).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_CATEGORY).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_SEX).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_EVENT_NAME).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_SUPECIFICATION).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_EVENT_TYPE_CODE).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_ROUNDS_COUNT).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_PERSONS_PER_GROUP).Value _
                            , (rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_OPERATION_WIND_GAUGE).Value <> 0) _
                            , (rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_MEASUREMENT_RT).Value <> 0) _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_MINUTE_PER_GROUP).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_CALL_START_TIMING).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_CALL_END_TIMING).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_TEMPLATE_FILE_NAME).Value _
                            , rowRange.Cells(1, CompetitionEventsColumnEnum.IDX_OUTPUT_FILE_NAME).Value _
                    )

                    Call m_Records.Add(.Self())
                End With
            End If
        Next rowRange
    End If

    ' Debug.Print m_Records.ToJson()
    Set ReadAllEvents = m_Records
End Function

Private  Function ThisSheet() As Worksheet
    Set ThisSheet = Sheets("Ží–Ú")
End Function
