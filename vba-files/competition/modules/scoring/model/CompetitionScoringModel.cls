VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CompetitionScoringModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/modules/scoring/model
Option Explicit
Implements IJsonConvertable
' TODO ê›íË

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function
