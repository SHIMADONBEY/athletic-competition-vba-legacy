VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CodeMasterRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''namespace=vba-files/competition/repository
Option Explicit

Private m_SexCode As Object
Private m_WindScalingCode As Object
Private m_ReactionTimeCode As Object
Private m_ResultUnitCode As Object
Private m_UsingScoreCode As Object
Private m_GroupStrategyCode As Object
Private m_OrderStrategyCode As Object
Private m_SameRankingStrategyCode As Object
Private m_EventCategoryCode As Object
Private m_EventMaster As EventMasterModels

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Get SexCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_SexCode Is Nothing) Then
        Set m_SexCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Sex").Rows()
            Call m_SexCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set SexCode = m_SexCode
End Property

Public Property Get WindScalingCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_WindScalingCode Is Nothing) Then
        Set m_WindScalingCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Wind").Rows()
            Call m_WindScalingCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set WindScalingCode = m_WindScalingCode
End Property

Public Property Get ReactionTimeCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_ReactionTimeCode Is Nothing) Then
        Set m_ReactionTimeCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Reaction").Rows()
            Call m_ReactionTimeCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set ReactionTimeCode = m_ReactionTimeCode
End Property

Public Property Get ResultUnitCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_ResultUnitCode Is Nothing) Then
        Set m_ResultUnitCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Result_Unit").Rows()
            Call m_ResultUnitCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set ResultUnitCode = m_ResultUnitCode
End Property

Public Property Get UsingScoreCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_UsingScoreCode Is Nothing) Then
        Set m_UsingScoreCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Point").Rows()
            Call m_UsingScoreCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set UsingScoreCode = m_UsingScoreCode
End Property

Public Property Get GroupStrategyCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_GroupStrategyCode Is Nothing) Then
        Set m_GroupStrategyCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Group_Strategy").Rows()
            Call m_GroupStrategyCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set GroupStrategyCode = m_GroupStrategyCode
End Property

Public Property Get OrderStrategyCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_OrderStrategyCode Is Nothing) Then
        Set m_OrderStrategyCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Lane_Strategy").Rows()
            Call m_OrderStrategyCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set OrderStrategyCode = m_OrderStrategyCode
End Property

Public Property Get SameRankingStrategyCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_SameRankingStrategyCode Is Nothing) Then
        Set m_SameRankingStrategyCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Same_Score").Rows()
            Call m_SameRankingStrategyCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set SameRankingStrategyCode = m_SameRankingStrategyCode
End Property

Public Property Get EventCategoryCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_EventCategoryCode Is Nothing) Then
        Set m_EventCategoryCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Category").Rows()
            Call m_EventCategoryCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set EventCategoryCode = m_EventCategoryCode
End Property

Public Property Get EventMaster(Optional Reload As Boolean = False) As EventMasterModels
    If (Reload Or m_EventMaster Is Nothing) Then
    ElseIf (m_EventMaster.Count() < 1) Then
    Else
        Set EventMaster = m_EventMaster
        Exit Property
    End If

    Dim vWindCache As Object: Set vWindCache = CreateObject("Scripting.Dictionary")
    Dim vWind As Variant
    For Each vWind In Me.WindScalingCode
        Call vWindCache.Add(Me.WindScalingCode.Item(vWind), vWind)
    Next vWind 

    Set m_EventMaster = New EventMasterModels
    Dim vRowRange As Range
    For Each vRowRange In Range("Code_Event").Rows()
        With New EventMasterModel
            Call .Initialize( _
                    vRowRange.Cells(1, 1).Value _
                    , vRowRange.Cells(1, 2).Value _
                    , vRowRange.Cells(1, 3).Value _
                    , vRowRange.Cells(1, 4).Value _
                    , vWindCache.Item(vRowRange.Cells(1, 5).Value) _
            )

            Call m_EventMaster.Add(.Self())
        End With
    Next vRowRange 

End Property

Public Sub Refresh()
    Dim vTemp As Object
    Set vTemp = Me.SexCode(True)
    Set vTemp = Me.WindScalingCode(True)
    Set vTemp = Me.ReactionTimeCode(True)
    Set vTemp = Me.ResultUnitCode(True)
    Set vTemp = Me.UsingScoreCode(True)
    Set vTemp = Me.GroupStrategyCode(True)
    Set vTemp = Me.OrderStrategyCode(True)
    Set vTemp = Me.SameRankingStrategyCode(True)
    Set vTemp = Me.EventCategoryCode(True)
    Set vTemp = Me.EventMaster(True)
    Set vTemp = Nothing
End Sub
