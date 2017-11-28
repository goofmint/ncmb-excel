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


Property Let Name(ByVal Name As String)
    m_Name = Name
End Property

Property Let ncmb(ByVal ncmb As clsNCMB)
    Set m_NCMB = ncmb
End Property

Public Function newData() As clsDataItem
    Set newData = New clsDataItem
    newData.dataStore = Me
    newData.ncmb = m_NCMB
End Function

Public Property Get Name() As String
    Name = m_Name
End Property

