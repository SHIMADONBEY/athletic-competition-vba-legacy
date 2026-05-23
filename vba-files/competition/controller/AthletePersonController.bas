Attribute VB_Name = "AthletePersonController"
'namespace=vba-files/competition/controller
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

Private Const HEADER_ROW As Long = 3

Private Function ThisSheet() As Worksheet
    Set ThisSheet = AthleteListSheet
End Function

Private Function RangeRepository() As Range
    Set RangeRepository = ThisSheet.Range("AthletePersonList")
End Function

Public Sub UpdateAll()
    Call AthletePersonRepository.ReadAllAthletePersons(True)
    Call AthletePersonRepository.ReadAllEntries(True)
End Sub

Public Sub ExportEntries()
    Dim vAllEntries As PersonalEntryModels: Set vAllEntries = AthletePersonRepository.ReadAllEntries(True)
    Dim vAllPersons As AthletePersonModels: Set vAllPersons = AthletePersonRepository.ReadAllAthletePersons(True)

    Dim vEventEntryGroup As Object: Set vEventEntryGroup = vAllEntries.GroupByEvent()
    Dim vAllCount As Variant: vAllCount = CDec(vAllEntries.Count())
    Dim vProcessCount As Long
    
    Dim vEvent As CompetitionEventModel
    For Each vEvent In CompetitionEventRepository.ReadAllEvents().All()
        If (vEventEntryGroup.Exists(vEvent.Id)) Then
            Dim vProgress As Variant: vProgress = Fix((CDec(vProcessCount * 1000) / vAllCount + CDec(5)) / CDec(10))
            Application.StatusBar = MessageFactory.Generate("SI012").Prompt(vEvent.EventKey(), vProgress)

            Dim vEntries As PersonalEntryModels: Set vEntries = vEventEntryGroup.Item(vEvent.Id)
            Call CompetitionEntriesLoader.LoadEntriesToFile(vEvent, vAllPersons, vEntries)

            vProcessCount = vProcessCount + vEntries.Count()
        End If
    Next vEvent

    Application.StatusBar = False
End Sub
