VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Name As String
Private m_NCMB As clsNCMB
Private m_Queries As Dictionary

Private Sub Class_Initialize()
    Set m_Queries = New Dictionary
End Sub

Property Let name(ByVal name As String)
    m_Name = name
End Property

Property Let ncmb(ByVal ncmb As clsNCMB)
    Set m_NCMB = ncmb
End Property

Public Function newData() As clsDataItem
    Set newData = New clsDataItem
    newData.dataStore = Me
    newData.ncmb = m_NCMB
End Function

Public Property Get name() As String
    name = m_Name
End Property

Public Sub makeDefault(Optional fieldName As String = "")
    If m_Queries.Exists("where") = False Then
        m_Queries.Add "where", New Dictionary
    End If
    If fieldName <> "" Then
        If m_Queries("where").Exists(fieldName) = False Then
            m_Queries("where").Add fieldName, New Dictionary
        End If
    End If
End Sub

Public Function equalTo(fieldName As String, value As Variant) As clsDataStore
    Call makeDefault
    m_Queries("where").Add fieldName, value
    Set equalTo = Me
End Function

Public Function greaterThan(fieldName As String, value As Integer) As clsDataStore
    Call makeDefault(fieldName)
    m_Queries("where")(fieldName).Add "$gt", value
    Set greaterThan = Me
End Function

Public Function greaterThanOrEqualTo(fieldName As String, value As Integer) As clsDataStore
    Call makeDefault(fieldName)
    m_Queries("where")(fieldName).Add "$gte", value
    Set greaterThan = Me
End Function

Public Function lessThan(fieldName As String, value As Integer) As clsDataStore
    Call makeDefault(fieldName)
    m_Queries("where")(fieldName).Add "$lt", value
    Set greaterThan = Me
End Function

Public Function lessThanOrEqualTo(fieldName As String, value As Integer) As clsDataStore
    Call makeDefault(fieldName)
    m_Queries("where")(fieldName).Add "$lte", value
    Set greaterThan = Me
End Function

Public Function notEqualTo(fieldName As String, value As Variant) As clsDataStore
    Call makeDefault(fieldName)
    m_Queries("where")(fieldName).Add "$ne", value
    Set greaterThan = Me
End Function



Public Function fetchAll() As clsDataItem()
    Dim dataItems() As clsDataItem
    
    Dim result As Dictionary
    Set result = m_NCMB.Request.GetRequest(m_Name, m_Queries)
    
    ReDim Preserve dataItems(result("results").Count)
    i = 0
    Dim params As Dictionary
    For Each params In result("results")
        Set params = params
        Set dataItems(i) = Me.newData
        Call dataItems(i).setFields(params)
        i = i + 1
    Next
    
    fetchAll = dataItems
End Function
