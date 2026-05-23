VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "EventMasterModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/model/event-master
Option Explicit
Implements IJsonConvertable

Private m_Records As Object
Private m_RecordIndex As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_RecordIndex = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As EventMasterModel)
    Call m_Records.Add(Item.Code, Item)
    Call m_RecordIndex.Add(Item.Name, Item.Code)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As EventMasterModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As EventMasterModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As EventMasterModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function FindByName(Name As String) As EventMasterModel
    If Not (m_RecordIndex.Exists(Name)) Then
        Set FindByName = Nothing
        Exit Function
    End If

    Dim vCode As String: vCode = m_RecordIndex.Item(Name)
    Set FindByName = m_Records.Item(vCode)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
