VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackRaceCloner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/modules/loader/cloner
Option Explicit
Implements ITemplateCloner

Private Const TEMPLATE_FILE As String = "\apps\track-race-vba.xlsm"

Private Sub ITemplateCloner_Clone(DestinationFile As String)
    Dim vTemplatePath As String: vTemplatePath = ThisWorkbook.Path & TEMPLATE_FILE
    If (Dir(vTemplatePath) = "") Then
        Err.Raise CustomErrorCodeEnum.TemplateFileNotFound, "TrackRaceCloner.Clone", MessageFactory.Generate("SE006").Prompt(vTemplatePath)
    End If
    FileCopy vTemplatePath, DestinationFile
End Sub
