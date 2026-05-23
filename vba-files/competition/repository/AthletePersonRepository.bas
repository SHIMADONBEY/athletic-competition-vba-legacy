Attribute VB_Name = "AthletePersonRepository"
'namespace=vba-files/competition/repository
Option Explicit
Option Private Module

Private Enum AthletePersonColumnEnum
    IDX_PERSONAL_CODE = 1
    IDX_SEX
    IDX_BIB
    IDX_PERSONAL_NAME
    IDX_AGE
    IDX_PERSONAL_PHONETIC
    IDX_PERSONAL_LATIN
    IDX_TEAM_NAME
    IDX_TEAM_PLACE
    IDX_PERSONAL_COUNTRY
    IDX_BIRTH_OF_DATE
    IDX_PERSONAL_GRADE
    IDX_ID
    IDX_TEAM_ID
End Enum

Private Enum EventEntryColumnEnum
    IDX_EVENT_NAME = 1
    IDX_DEMO_ENTRY
    IDX_EVENT_ID
    IDX_QUALIFIED_1
    IDX_QUALIFIED_2
    IDX_QUALIFIED_3
End Enum

Private Const EVENT_ENTRY_COLUMN_COUNT As Long = 6
Private Const EVENT_ENTRY_COLUMN_OFFSET As Long = 14
Private Const ENTRIES_PER_PERSON_MAX_COUNT As Long = 10

Private m_PersonRepository As AthletePersonModels
Private m_EntryRepository As PersonalEntryModels

Private Function ThisSheet() As Worksheet
    Set ThisSheet = AthleteListSheet
End Function

Private Function RangeRepository() As Range
    Set RangeRepository = ThisSheet.Range("AthletePersonList")
End Function

Public Function ReadAllAthletePersons(Optional Reload As Boolean = False) As AthletePersonModels 
    If (Reload Or m_PersonRepository Is Nothing) Then
        Dim vRecords As AthletePersonModels: Set vRecords = New AthletePersonModels

        Dim vRowRange As Range
        For Each vRowRange In RangeRepository.Rows()
            Dim vRecord As AthletePersonModel: Set vRecord = ReadPersonFromRow(vRowRange)
            If (vRecord Is Nothing) Then 
                ' DO NOTHING
            Else
                Call vRecords.Add(vRecord)
                vRowRange.Cells(1, AthletePersonColumnEnum.IDX_ID).Value = vRecord.Id
            End If
        Next vRowRange

        Set m_PersonRepository = vRecords
    End If

    ' Debug.Print m_PersonRepository.ToJson()
    Set ReadAllAthletePersons = m_PersonRepository
End Function

Public Function ReadAllEntries(Optional Reload As Boolean = False) As PersonalEntryModels
    If (Reload Or m_EntryRepository Is Nothing) Then
        Dim vRecords As PersonalEntryModels: Set vRecords = New PersonalEntryModels
        Dim vRowRange As Range
        For Each vRowRange In RangeRepository.Rows()
            Call vRecords.AddRange(ReadEntriesFromRow(vRowRange))
        Next vRowRange 
        Set m_EntryRepository = vRecords
    End If

    ' Debug.Print m_EntryRepository.ToJson()
    Set ReadAllEntries = m_EntryRepository
End Function

Private Function ReadPersonFromRow(RowRange As Range) As AthletePersonModel
    Dim pSex As String          : pSex = RowRange.Cells(1, AthletePersonColumnEnum.IDX_SEX).Value
    Dim pName As String         : pName = RowRange.Cells(1, AthletePersonColumnEnum.IDX_PERSONAL_NAME).Value
    Dim pTeamName As String     : pTeamName = RowRange.Cells(1, AthletePersonColumnEnum.IDX_TEAM_NAME).Value

    If (pName = "" Or pSex = "" Or pTeamName = "") Then 
        Set ReadPersonFromRow = Nothing
        Exit Function
    End If

    With New AthletePersonModel
        Call .Initialize( _
                pSex _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_BIB).Value _
                , pName _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_AGE).Value _
                , pTeamName _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_PERSONAL_PHONETIC).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_PERSONAL_LATIN).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_TEAM_PLACE).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_PERSONAL_COUNTRY).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_PERSONAL_GRADE).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_BIRTH_OF_DATE).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_PERSONAL_CODE).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_TEAM_ID).Value _
                , RowRange.Cells(1, AthletePersonColumnEnum.IDX_ID).Value _
        )

        Set ReadPersonFromRow = .Self()
    End With
End Function

Private Function ReadEntriesFromRow(RowRange As Range) As Collection
    Dim vEntries As Collection              : Set vEntries = New Collection
    Dim pPersonId As String                 : pPersonId = RowRange.Cells(1, AthletePersonColumnEnum.IDX_ID).Value
    Dim vMinuteDelim As String              : vMinuteDelim = CompetitionConfigFactory.GetConfigInstance().MinuteDelimiter
    Dim vSecondDelim As String              : vSecondDelim = CompetitionConfigFactory.GetConfigInstance().SecondDelimiter
    Dim vMetreDelim As String               : vMetreDelim = CompetitionConfigFactory.GetConfigInstance().MeterDelimiter
    Dim vEvents As CompetitionEventModels   : Set vEvents =  CompetitionEventRepository.ReadAllEvents()

    If (pPersonId = "") Then
        Set ReadEntriesFromRow = vEntries
        Exit Function
    End If

    Dim pEventRange As Range: Set pEventRange = RowRange.Cells(1, EVENT_ENTRY_COLUMN_OFFSET + 1).Resize(1, EVENT_ENTRY_COLUMN_COUNT)
    Dim i As Long
    For i = 1 To CompetitionConfigFactory.GetConfigInstance().CompetibleEventsCount
        Dim pEventName As String            : pEventName = pEventRange.Cells(1, EventEntryColumnEnum.IDX_EVENT_NAME).Value
        Dim vEvent As CompetitionEventModel : Set vEvent = vEvents.FindByEvent(pEventName)
        If (pEventName = "") Then 
            pEventRange.Cells(1, EventEntryColumnEnum.IDX_EVENT_ID).Value = Empty
        ElseIf(vEvent Is Nothing) Then
            pEventRange.Cells(1, EventEntryColumnEnum.IDX_EVENT_ID).Value = Empty
        Else
            With New PersonalEntryModel
                Call .Initialize( _
                        pPersonId _
                        , i _
                        , vEvent.Id _
                        , Not (pEventRange.Cells(1, EventEntryColumnEnum.IDX_DEMO_ENTRY).Value = "") _
                        , ResultValueParser.Parse(pEventRange.Cells(1, EventEntryColumnEnum.IDX_QUALIFIED_1).Value, vEvent.EventType, vMinuteDelim, vSecondDelim, vMetreDelim) _
                        , ResultValueParser.Parse(pEventRange.Cells(1, EventEntryColumnEnum.IDX_QUALIFIED_2).Value, vEvent.EventType, vMinuteDelim, vSecondDelim, vMetreDelim) _
                        , ResultValueParser.Parse(pEventRange.Cells(1, EventEntryColumnEnum.IDX_QUALIFIED_3).Value, vEvent.EventType, vMinuteDelim, vSecondDelim, vMetreDelim) _
                )

                Call vEntries.Add(.Self())
                pEventRange.Cells(1, EventEntryColumnEnum.IDX_EVENT_ID).Value = vEvent.Id
            End With
        End If

        If (i + 1 > ENTRIES_PER_PERSON_MAX_COUNT) Then
            ' “ü—Í—“‚ª‚È‚¢‚½‚ß‚±‚±‚ÅI—¹.
            Exit For
        End If

        Set pEventRange = pEventRange.Offset(0, EVENT_ENTRY_COLUMN_COUNT)
    Next i

    Set ReadEntriesFromRow = vEntries
End Function
