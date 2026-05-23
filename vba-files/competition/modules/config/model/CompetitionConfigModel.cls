VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CompetitionConfigModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/modules/config/model
Option Explicit
Implements IJsonConvertable

Private m_BasicInfo As CompetitionInfoModel
Private m_Sponsors As SponsorModels
Private m_ScoringConfig As CompetitionScoringModel
Private m_CompetibleEventsCount As Long
Private m_LanesCount As Long
Private m_MinuteDelimiter As String
Private m_SecondDelimiter As String
Private m_MeterDelimiter As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
            BasicInfo As CompetitionInfoModel _
            , Sponsors As SponsorModels _
            , ScoringConfig As CompetitionScoringModel _
            , CompetibleEventsCount As Long _
            , LanesCount As Long _
            , MinuteDelimiter As String _
            , SecondDelimiter As String _
            , MeterDelimiter As String)

    If m_Initialized Then Err.Raise(91)
    Set m_BasicInfo = BasicInfo
    Set m_Sponsors = Sponsors
    Set m_ScoringConfig = ScoringConfig
    m_CompetibleEventsCount = CompetibleEventsCount
    m_LanesCount = LanesCount
    m_MinuteDelimiter = MinuteDelimiter
    m_SecondDelimiter = SecondDelimiter
    m_MeterDelimiter = MeterDelimiter
    m_Initialized = True
End Sub

Public Property Get BasicInfo() As CompetitionInfoModel
    If Not m_Initialized Then Err.Raise(91)
    Set BasicInfo = m_BasicInfo
End Property

Public Property Get Sponsors() As SponsorModels
    If Not m_Initialized Then Err.Raise(91)
    Set Sponsors = m_Sponsors
End Property

Public Property Get ScoringConfig() As CompetitionScoringModel
    If Not m_Initialized Then Err.Raise(91)
    Set ScoringConfig = m_ScoringConfig
End Property

Public Property Get CompetibleEventsCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    CompetibleEventsCount = m_CompetibleEventsCount
End Property

Public Property Get LanesCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    LanesCount = m_LanesCount
End Property

Public Property Get MinuteDelimiter() As String
    If Not m_Initialized Then Err.Raise(91)
    MinuteDelimiter = m_MinuteDelimiter
End Property

Public Property Get SecondDelimiter() As String
    If Not m_Initialized Then Err.Raise(91)
    SecondDelimiter = m_SecondDelimiter
End Property

Public Property Get MeterDelimiter() As String
    If Not m_Initialized Then Err.Raise(91)
    MeterDelimiter = m_MeterDelimiter
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("basic_info", Me.BasicInfo)
    Call jsonDictionary.Add("sponsors", Me.Sponsors)
    Call jsonDictionary.Add("scoring_config", Me.ScoringConfig)
    Call jsonDictionary.Add("competible_events_count", Me.CompetibleEventsCount)
    Call jsonDictionary.Add("lanes_count", Me.LanesCount)
    Call jsonDictionary.Add("minute_delimiter", Me.MinuteDelimiter)
    Call jsonDictionary.Add("second_delimiter", Me.SecondDelimiter)
    Call jsonDictionary.Add("meter_delimiter", Me.MeterDelimiter)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
