VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "EventMasterModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/model/event-master
Option Explicit
Implements IJsonConvertable

Private m_Code As String
Private m_Name As String
Private m_LatinName As String
Private m_Category As String
Private m_NeedWindMeasuring As Integer

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(Code As String, Name As String, LatinName As String, Category As String, NeedWindMeasuring As Integer)
    If m_Initialized Then Err.Raise(91)
    m_Code = Code
    m_Name = Name
    m_LatinName = LatinName
    m_Category = Category
    m_NeedWindMeasuring = NeedWindMeasuring
    m_Initialized = True
End Sub

Public Property Get Code() As String
    If Not m_Initialized Then Err.Raise(91)
    Code = m_Code
End Property

Public Property Get Name() As String
    If Not m_Initialized Then Err.Raise(91)
    Name = m_Name
End Property

Public Property Get LatinName() As String
    If Not m_Initialized Then Err.Raise(91)
    LatinName = m_LatinName
End Property

Public Property Get Category() As String
    If Not m_Initialized Then Err.Raise(91)
    Category = m_Category
End Property

Public Property Get NeedWindMeasuring() As Integer
    If Not m_Initialized Then Err.Raise(91)
    NeedWindMeasuring = m_NeedWindMeasuring
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("code", Me.Code)
    Call jsonDictionary.Add("name", Me.Name)
    Call jsonDictionary.Add("latin_name", Me.LatinName)
    Call jsonDictionary.Add("category", Me.Category)
    Call jsonDictionary.Add("need_wind_measuring", Me.NeedWindMeasuring)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
