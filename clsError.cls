VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Code As String
Private m_Message As String

Private Sub Class_Initialize()

End Sub


Public Property Let Code(strValue As String)
    m_Code = strValue
End Property

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Message(strValue As String)
    m_Message = strValue
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

