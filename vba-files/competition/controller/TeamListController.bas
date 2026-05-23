Attribute VB_Name = "TeamListController"
'namespace=vba-files/competition/controller
Option Explicit
Option Private Module

Private Enum TeamListColumnEnum
    IDX_NAME = 1
    IDX_PHONETIC
    IDX_LATIN_NAME
    IDX_FULL_NAME
    IDX_PLACE
    IDX_COUNTRY
    IDX_GRADES
    IDX_LEADER_NAME
    IDX_LEADER_CONTACT_1
    IDX_LEADER_CONTACT_2
    IDX_PARTICIPATING_INPUT
    IDX_ENTRY_COLUMN
    IDX_PARTICIPATING
    IDX_ID
End Enum

Private Const HEADER_ROW As Long = 3

Private Function ThisSheet() As Worksheet
    Set ThisSheet = TeamListSheet
End Function

Public Sub UpdateTeamList()
    Call TeamRepository.ReadAllTeams(True)
End Sub
