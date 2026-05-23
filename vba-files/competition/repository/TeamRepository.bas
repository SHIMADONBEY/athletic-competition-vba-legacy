Attribute VB_Name = "TeamRepository"
'namespace=vba-files/competition/repoistory
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

Private m_Repository As TeamModels

Private Function ThisSheet() As Worksheet
    Set ThisSheet = TeamListSheet
End Function

Private Function RangeRepository() As Range
    Set RangeRepository = ThisSheet.Range("TeamList")
End Function

Public Function ReadAllTeams(Optional Reload As Boolean = False) As TeamModels
    If (Reload Or m_Repository Is Nothing) Then
        Dim vRecords As TeamModels: Set vRecords = New TeamModels

        Dim vRowRange As Range
        For Each vRowRange In RangeRepository.Rows()
            Dim vRecord As TeamModel: Set vRecord = ReadFromRow(vRowRange)
            If (vRecord Is Nothing) Then 
                ' DO NOTHING
            Else
                Call vRecords.Add(vRecord)
                vRowRange.Cells(1, TeamListColumnEnum.IDX_ID).Value = vRecord.Id
            End If
        Next vRowRange 

        Set m_Repository = vRecords
    End If

    ' Debug.Print m_Repository.ToJson()
    Set ReadAllTeams = m_Repository
End Function

Private Function ReadFromRow(RowRange As Range) As TeamModel
    Dim vTeamName As String: vTeamName = RowRange.Cells(1, TeamListColumnEnum.IDX_NAME).Value

    If (vTeamName = "") Then
        Set ReadFromRow = Nothing
        Exit Function
    End If

    Dim vContacts As Collection: Set vContacts = New Collection
    Call vContacts.Add(RowRange.Cells(1, TeamListColumnEnum.IDX_LEADER_CONTACT_1).Value)
    Call vContacts.Add(RowRange.Cells(1, TeamListColumnEnum.IDX_LEADER_CONTACT_2).Value)

    With New TeamModel
        Call .Initialize( _
                vTeamName _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_PHONETIC).Value _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_LATIN_NAME).Value _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_FULL_NAME).Value _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_PLACE).Value _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_COUNTRY).Value _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_GRADES).Value _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_LEADER_NAME).Value _
                , vContacts _
                , (RowRange.Cells(1, TeamListColumnEnum.IDX_PARTICIPATING).Value = "") _
                , RowRange.Cells(1, TeamListColumnEnum.IDX_ID).Value _
        )

        Set ReadFromRow = .Self()
    End With
End Function
