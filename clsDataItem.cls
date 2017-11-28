VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Fields As Dictionary
Private m_NCMB As clsNCMB
Private m_DataStore As clsDataStore
Private m_Error As clsError

Private Sub Class_Initialize()
    Set m_Fields = New Dictionary
End Sub

Public Sub Field(Name As String, value As Variant)
    m_Fields.Add Name, value
End Sub

Property Let ncmb(ByVal ncmb As clsNCMB)
    Set m_NCMB = ncmb
End Property

Property Let dataStore(ByVal dataStore As clsDataStore)
    Set m_DataStore = dataStore
End Property


Public Function Save() As Boolean
    Dim result As Dictionary
    If m_Fields.Exists("objectId") Then
        Set result = m_NCMB.Request.PutRequest(m_DataStore.Name, m_Fields)
    Else
        Set result = m_NCMB.Request.PostRequest(m_DataStore.Name, m_Fields)
    End If
    
    ' èàóùåãâ 
    If result.Exists("objectId") Then
        Save = True
        For Each key In result
            m_Fields.Add key, result(key)
        Next
    Else
        Save = False
        m_Error = New clsError
        m_Error.Code = result("code")
        m_Error.Message = result("error")
    End If
End Function

Public Property Get Error() As String
    Error = m_Error
End Property

