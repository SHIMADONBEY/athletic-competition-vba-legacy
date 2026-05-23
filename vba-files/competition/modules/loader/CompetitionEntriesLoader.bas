Attribute VB_Name = "CompetitionEntriesLoader"
'namespace=vba-files/competition/modules/loader
Option Explicit
Option Private Module

Public Sub LoadEntriesToFile(CompetitionEvent As CompetitionEventModel, Persons As AthletePersonModels, Entries As PersonalEntryModels)
    Dim vFileName As String: vFileName = ThisWorkbook.Path & "\" & CompetitionEvent.EventKey() & ".xlsm"
    If (Dir(vFileName, 0) = "") Then
        Err.Raise CustomErrorCodeEnum.EventFileNotFound, "CompetitionRoundLoader.LoadCompetitionRoundToFile", MessageFactory.Generate("SE009").Prompt(CompetitionEvent.EventKey())
    End If

    Dim vFilteredPersons As AthletePersonModels: Set vFilteredPersons = New AthletePersonModels
    Dim vEntry As PersonalEntryModel
    For Each vEntry In Entries.All()
        Call vFilteredPersons.Add(Persons.Item(vEntry.PersonalId))
    Next vEntry 

    Dim vCurrentBook As Workbook: Set vCurrentBook = Workbooks.Open(vFileName)
    Application.Run "'" & vFileName & "'!EntryListImporter.ImportEntries", Entries.ToJson(), vFilteredPersons.ToJson()
    Call vCurrentBook.Close(True)
    Set vCurrentBook = Nothing
End Sub
