VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNCMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Url As String
Private m_Version As String
Private m_Fqdn As String

Private m_ScriptVersion As String
Private m_ScriptFqdn As String
Private m_Port As Integer
Private m_Protocol As String

Private m_SignatureMethod As String
Private m_SignatureVersion As Integer

Private m_ApplicationKey As String
Private m_ClientKey As String
Private m_Request As clsRequest

Private Sub Class_Initialize()
    m_Version = "2013-09-01"
    m_Fqdn = "mb.api.cloud.nifty.com"
    m_ScriptVersion = "2015-09-01"
    m_ScriptFqdn = "script.mb.api.cloud.nifty.com"
    m_Port = 443
    m_Protocol = "https:"
    m_SignatureMethod = "HmacSHA256"
    m_SignatureVersion = 2
    Set m_Request = New clsRequest
    m_Request.ncmb = Me
End Sub


Public Sub setKeys(ApplicationKey As String, ClientKey As String)
    m_ApplicationKey = ApplicationKey
    m_ClientKey = ClientKey
End Sub


Property Let ApplicationKey(ByVal ApplicationKey As String)
    m_ApplicationKey = ApplicationKey
End Property

Property Let ClientKey(ByVal ClientKey As String)
    m_ClientKey = ClientKey
End Property

Public Function dataStore(Name As String) As clsDataStore
    Set dataStore = New clsDataStore
    dataStore.ncmb = Me
    dataStore.Name = Name
End Function

Public Property Get Version() As String
    Version = m_Version
End Property

Public Property Get Fqdn() As String
    Fqdn = m_Fqdn
End Property

Public Property Get Protocol() As String
    Protocol = m_Protocol
End Property

Public Property Get SignatureMethod() As String
    SignatureMethod = m_SignatureMethod
End Property

Public Property Get SignatureVersion() As String
    SignatureVersion = m_SignatureVersion
End Property

Public Property Get ApplicationKey() As String
    ApplicationKey = m_ApplicationKey
End Property

Public Property Get ClientKey() As String
    ClientKey = m_ClientKey
End Property

Public Property Get Request() As clsRequest
    Set Request = m_Request
End Property

