Attribute VB_Name = "TemplateClonerFactory"
'namespace=vba-files/competition/modules/loader/cloner
Option Explicit
Option Private Module

Public Function GenerateTemplateCloner(EventType As String) As ITemplateCloner
    Select Case EventType
        Case EventTypeConstants.TRACK
            Set GenerateTemplateCloner = New TrackRaceCloner
        Case Else 
            Set GenerateTemplateCloner = Nothing
    End Select
End Function
