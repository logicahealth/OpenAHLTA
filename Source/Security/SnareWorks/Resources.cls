VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Resources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_Resources

Private Const CLASS_NAME As String = "CHCSII_SnareWorks.Resources"
Private m_iCount As Integer
Private m_objResources As New Collection

Friend Function AddResource(ByVal Name As String, ByVal UserID As String, ByVal Password As String, ResourceData As Variant)
    'ResourceData could be a string or a Collection of Key/Value Pairs
    Dim colAttributes As New Collection
    Dim sResourceData As String
    Dim objKeyValuePair As KeyValuePair
    Dim i As Integer
'<< sCR 12110
    On Error Resume Next
    If (TypeOf ResourceData Is Collection) Then
        For i = 1 To ResourceData.Count
            colAttributes.Add ResourceData.Item(i), ResourceData.Item(i).Key
        Next
        sResourceData = vbNullString
    Else
        Set objKeyValuePair = New KeyValuePair
        objKeyValuePair.Key = "ResourceData"
        objKeyValuePair.Value = ResourceData
        colAttributes.Add objKeyValuePair, "ResourceData"
        sResourceData = ResourceData
    End If
    
    Dim objResource As New Resource
    
    With objResource
        .Name = Name
        .UserID = UserID
        .Password = Password
        .ResourceData = sResourceData
        Set .Attributes = colAttributes
        m_objResources.Add objResource, .Name
    End With
End Function

Friend Function Item(Index As Variant) As Resource
    Set Item = m_objResources.Item(Index)
End Function

Private Property Get ICHCSII_Resources_Count() As Integer
    ICHCSII_Resources_Count = m_objResources.Count
End Property

Private Function ICHCSII_Resources_Item(Index As Variant) As CHCSII_ClientSecurity.ICHCSII_Resource
    On Error GoTo errhandler
    Set ICHCSII_Resources_Item = m_objResources.Item(Index)
    
    Exit Function

errhandler:
    err.Raise SEC_NO_USER_RESOURCE, CLASS_NAME, GetErrMsg(SEC_NO_USER_RESOURCE)

End Function

