VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_NCMB As clsNCMB

Private Sub Class_Initialize()
End Sub

Property Let ncmb(ByVal ncmb As clsNCMB)
    Set m_NCMB = ncmb
End Property

Public Function login(userName As String, password As String) As clsDataItem
    Dim params As Dictionary
    Set params = New Dictionary
    params.Add "userName", userName
    params.Add "password", password
    Dim userClass As clsDataStore
    Set userClass = m_NCMB.dataStore("user")
    Set login = userClass.newData
    
    Dim result As Dictionary
    Set result = m_NCMB.Request.GetRequest("login", params)
    login.setFields result
    m_NCMB.SessionToken = result("sessionToken")
End Function

