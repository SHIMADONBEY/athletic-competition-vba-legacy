VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "SponsorModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/modules/config/model
Option Explicit
Implements IJsonConvertable

Private m_Category As String
Private m_Sponsor As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize(Category As String, Sponsor As String)
    If m_Initialized Then Err.Raise(91)
    m_Category = Category
    m_Sponsor = Sponsor
    m_Initialized = True
End Sub

Public Property Get Category() As String
    If Not m_Initialized Then Err.Raise(91)
    Category = m_Category
End Property

Public Property Get Sponsor() As String
    If Not m_Initialized Then Err.Raise(91)
    Sponsor = m_Sponsor
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("category", Me.Category)
    Call jsonDictionary.Add("sponsor", Me.Sponsor)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
