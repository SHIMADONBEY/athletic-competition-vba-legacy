Attribute VB_Name = "GameRecordLoader"
'namespace=vba-files/competition/modules/loader
Option Explicit
Option Private Module

Public Sub LoadGameRecordsToFile(CompetitionEvent As CompetitionEventModel, GameRecords As GameRecordModels)
    Dim vFileName As String: vFileName = ThisWorkbook.Path & "\" & CompetitionEvent.EventKey() & ".xlsm"

    If (Dir(vFileName, 0) = "") Then
        Err.Raise CustomErrorCodeEnum.EventFileNotFound, "GameRecordLoader.LoadGameRecordsToFile", MessageFactory.Generate("SE009").Prompt(CompetitionEvent.EventKey())
    End If

    Dim vCurrentBook As Workbook: Set vCurrentBook = Workbooks.Open(vFileName)
    Application.Run "'" & vFileName & "'!GameRecordsImporter.ImportGameRecords", CompetitionEvent.ToJson(), GameRecords.ToJson()
    Call vCurrentBook.Close(true)
    Set vCurrentBook = Nothing
End Sub
