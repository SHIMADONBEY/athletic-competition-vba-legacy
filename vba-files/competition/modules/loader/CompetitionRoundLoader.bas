Attribute VB_Name = "CompetitionRoundLoader"
'namespace=vba-files/competitionmodules/loader
Option Explicit
Option Private Module

Public Sub LoadCompetitionRoundToFile(CompetitionEvent As CompetitionEventModel, CompetitionRounds As CompetitionRoundModels)
    Dim vFileName As String: vFileName = ThisWorkbook.Path & "\" & CompetitionEvent.EventKey() & ".xlsm"
    If (Dir(vFileName, 0) = "") Then
        Err.Raise CustomErrorCodeEnum.EventFileNotFound, "CompetitionRoundLoader.LoadCompetitionRoundToFile", MessageFactory.Generate("SE009").Prompt(CompetitionEvent.EventKey())
    End If

    Dim vCurrentBook As Workbook: Set vCurrentBook = Workbooks.Open(vFileName)
    Application.Run "'" & vFileName & "'!CompetitionRoundImporter.ImportCompetitionRounds", CompetitionEvent.ToJson(), CompetitionRounds.ToJson()
    Call vCurrentBook.Close(True)
    Set vCurrentBook = Nothing
End Sub