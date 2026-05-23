VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "SponsorModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/modules/config/model
Option Explicit
Implements IJsonConvertable

Private m_Records As Collection

Private Sub Class_Initialize()
    Set m_Records = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As SponsorModel)
    Call m_Records.Add(Item, Item.Category)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As SponsorModel
    For Each recordItem In Items
        Call m_Records.Add(recordItem)
    Next recordItem 
End Sub

Public  Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordItem As SponsorModel    

    For Each recordItem In m_Records
        Call recordItems.Add(recordItem)
    Next recordItem 

    Set All = recordItems
End Function

Public Function FindByCategory(Category As String) As SponsorModel
    Set FindByCategory = m_Records.Item(Category)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All())
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
