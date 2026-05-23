Attribute VB_Name = "CompetitionSettingController"
'namespace=vba-files/competition/controller
Option Explicit
Option Private Module

Public Sub UpdateCompetition()
    Call CompetitionConfigFactory.UpdateConfig()
End Sub

Private Function ThisSheet() As Worksheet
    Set ThisSheet = Sheets("ëÂâÔê›íË")
End Function
