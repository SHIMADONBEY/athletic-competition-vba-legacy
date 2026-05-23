VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "TeamModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/model/team
Option Explicit
Implements IJsonConvertable

Private m_Name As String
Private m_Phonetic As String
Private m_LatinName As String
Private m_FullName As String
Private m_Place As String
Private m_Country As String
Private m_Grade As String
Private m_LeaderName As String
Private m_LeaderContacts As Collection
Private m_Participation As Boolean
Private m_Id As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
        Name As String _
        , Phonetic As String _
        , LatinName As String _
        , FullName As String _
        , Place As String _
        , Country As String _
        , Grade As String _
        , LeaderName As String _
        , LeaderContacts As Collection _
        , Participation As Boolean _
        , Optional Id As String = "" _
    )
    If m_Initialized Then Err.Raise(91)
    m_Name = Name
    m_Phonetic = Phonetic
    m_LatinName = LatinName
    m_FullName = FullName
    m_Place = Place
    m_Country = Country
    m_Grade = Grade
    m_LeaderName = LeaderName
    Set m_LeaderContacts = LeaderContacts
    m_Participation = Participation

    If (Id = "") Then 
        m_Id = UuidFactory.GenerateUuid()
    Else
        m_Id = Id
    End If

    m_Initialized = True
End Sub

Public Property Get Name() As String
    If Not m_Initialized Then Err.Raise(91)
    Name = m_Name
End Property

Public Property Get Phonetic() As String
    If Not m_Initialized Then Err.Raise(91)
    Phonetic = m_Phonetic
End Property

Public Property Get LatinName() As String
    If Not m_Initialized Then Err.Raise(91)
    LatinName = m_LatinName
End Property

Public Property Get FullName() As String
    If Not m_Initialized Then Err.Raise(91)
    FullName = m_FullName
End Property

Public Property Get Place() As String
    If Not m_Initialized Then Err.Raise(91)
    Place = m_Place
End Property

Public Property Get Country() As String
    If Not m_Initialized Then Err.Raise(91)
    Country = m_Country
End Property

Public Property Get Grade() As String
    If Not m_Initialized Then Err.Raise(91)
    Grade = m_Grade
End Property

Public Property Get LeaderName() As String
    If Not m_Initialized Then Err.Raise(91)
    LeaderName = m_LeaderName
End Property

Public Property Get LeaderContacts() As Collection
    If Not m_Initialized Then Err.Raise(91)
    Set LeaderContacts = m_LeaderContacts
End Property

Public Property Get Participation() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    Participation = m_Participation
End Property

Public Property Get Id() As String
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("name", Me.Name)
    Call jsonDictionary.Add("phonetic", Me.Phonetic)
    Call jsonDictionary.Add("latin_name", Me.LatinName)
    Call jsonDictionary.Add("full_name", Me.FullName)
    Call jsonDictionary.Add("place", Me.Place)
    Call jsonDictionary.Add("country", Me.Country)
    Call jsonDictionary.Add("grade", Me.Grade)
    Call jsonDictionary.Add("leader_name", Me.LeaderName)
    Call jsonDictionary.Add("leader_contacts", Me.LeaderContacts)
    Call jsonDictionary.Add("participation", Me.Participation)
    Call jsonDictionary.Add("id", Me.Id)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
