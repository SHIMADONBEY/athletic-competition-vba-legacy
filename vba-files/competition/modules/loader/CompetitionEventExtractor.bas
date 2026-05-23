Attribute VB_Name = "CompetitionEventExtractor"
'namespace=vba-files/competitionmodules/loader
Option Explicit
Option Private Module

Public Sub ExtractCompetionEventToFile(CompetitionSettings As CompetitionConfigModel, CompetitionEvent As CompetitionEventModel)
    Dim vFileName As String: vFileName = ThisWorkbook.Path & "\" & CompetitionEvent.EventKey() & ".xlsm"
    If (Dir(vFileName, 0) <> "") Then
        If Not (MessageFactory.Generate("SQ003").ToMsgBox(CompetitionEvent.EventKey()) = vbYes) Then
            Exit Sub
        End If
        Kill vFileName
    End If

    Dim vCompetitionEvent As CompetitionEventModel
    Dim vTemplateCloner As ITemplateCloner: Set vTemplateCloner = TemplateClonerFactory.GenerateTemplateCloner(CompetitionEvent.EventType)
    If (vTemplateCloner Is Nothing) Then
        Err.Raise CustomErrorCodeEnum.NotApplicated, "CompetitionEventExtractor.ExtractCompetionEventToFile", MessageFactory.Generate("SE005").Prompt(CompetitionEvent.EventKey())
    End If

    Call vTemplateCloner.Clone(vFileName)

    Dim vCurrentBook As Workbook: Set vCurrentBook = Workbooks.Open(vFileName)
    Application.Run "'" & vFileName & "'!ImportCompetitionEvent", CompetitionSettings.BasicInfo.ToJson(), CompetitionEvent.ToJson(), CompetitionSettings.ToJson()
    Call vCurrentBook.Close(True)
    Set vCurrentBook = Nothing
End Sub
