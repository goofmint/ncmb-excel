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

Public Sub Field(name As String, value As Variant)
    If m_Fields.Exists(name) Then
        m_Fields(name) = value
    Else
        m_Fields.Add name, value
    End If
End Sub

Public Sub setFields(fields As Dictionary)
    For Each key In fields
        Field key, fields(key)
    Next
End Sub

Public Function val(name As String) As Variant
    val = m_Fields(name)
End Function

Property Let ncmb(ByVal ncmb As clsNCMB)
    Set m_NCMB = ncmb
End Property

Property Let dataStore(ByVal dataStore As clsDataStore)
    Set m_DataStore = dataStore
End Property


Public Function Save() As Boolean
    Dim result As Dictionary
    Dim fields As Dictionary
    Set fields = m_Fields
    Debug.Print Join(fields.Keys, ",")
    If m_Fields.Exists("objectId") Then
        If fields.Exists("createDate") Then
            fields.Remove ("createDate")
        End If
        If fields.Exists("updateDate") Then
            fields.Remove ("updateDate")
        End If
        Set result = m_NCMB.Request.PutRequest(m_DataStore.name, m_Fields)
    Else
        Set result = m_NCMB.Request.PostRequest(m_DataStore.name, m_Fields)
    End If
    
    ' èàóùåãâ 
    If result.Exists("code") Then
        Save = False
        ' {"code":"E404005","error":"No such application."}
        m_Error = New clsError
        m_Error.Code = result("code")
        m_Error.Message = result("error")
    Else
        Save = True
        For Each key In result
            m_Fields.Add key, result(key)
        Next
    End If
End Function

Public Property Get Error() As String
    Error = m_Error
End Property

