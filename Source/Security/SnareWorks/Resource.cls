VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_Resource

Private m_sName As String
Private m_sPassword As String
Private m_sUserId As String
Private m_sResourceData As String
Private m_colAttributes As Collection


Friend Property Let Name(ByVal Name As String)
    m_sName = Name
End Property

Friend Property Get Name() As String
    Name = m_sName
End Property

Friend Property Let Password(ByVal Password As String)
    m_sPassword = Password
End Property

Friend Property Let UserID(ByVal UserID As String)
    m_sUserId = UserID
End Property

Friend Property Let ResourceData(ByVal ResourceData As String)
    m_sResourceData = ResourceData
End Property

Friend Property Get ResourceData() As String
    ResourceData = m_sResourceData
End Property

Friend Property Set Attributes(colAttributes As Collection)
    Set m_colAttributes = colAttributes
End Property

Friend Property Get Attributes() As Collection
    Set Attributes = m_colAttributes
End Property

Private Function ICHCSII_Resource_GetAttribute(AttributeName As String) As String
    On Error GoTo errhandler
    ICHCSII_Resource_GetAttribute = m_colAttributes.Item(AttributeName).Value
errhandler:
End Function

Private Function ICHCSII_Resource_GetAttributeNames() As Collection
    Dim colAttributeNames As New Collection
    Dim i As Integer
    
    For i = 1 To m_colAttributes.Count
        colAttributeNames.Add m_colAttributes(i).Key
    Next
    
    Set ICHCSII_Resource_GetAttributeNames = colAttributeNames
End Function

Private Property Get ICHCSII_Resource_Name() As String
    ICHCSII_Resource_Name = m_sName
End Property

Private Property Get ICHCSII_Resource_Password() As String
    ICHCSII_Resource_Password = m_sPassword
End Property

Private Property Get ICHCSII_Resource_ResourceData() As String
    ICHCSII_Resource_ResourceData = m_sResourceData
End Property

Private Property Get ICHCSII_Resource_UserId() As String
    ICHCSII_Resource_UserId = m_sUserId
End Property
