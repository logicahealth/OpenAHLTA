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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICHCSII_Resources

Private Const CLASS_NAME As String = "C2T_Security.Resources"
Private m_iCount As Integer
Private m_objResources As New Collection

Friend Function AddResource(ByVal Name As String, _
                            ByVal UserId As String, _
                            ByVal Password As String, _
                            ResourceData As Variant, _
                   Optional Attributes As Collection)
    'ResourceData could be a string or a Collection of Key/Value Pairs
    Dim colAttributes As New Collection
    Dim sResourceData As String
    Dim objKeyValuePair As KeyValuePair
    Dim i As Integer
    
    If (Not (Attributes Is Nothing)) Then
        Set colAttributes = Attributes
        sResourceData = ResourceData
    Else
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
    End If
    
    
    Dim objResource As New Resource
    With objResource
        .Name = Name
        .UserId = UserId
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
    On Error Resume Next
    Set ICHCSII_Resources_Item = m_objResources.Item(Index)
    If Err Then
        On Error GoTo 0
        Err.Raise SEC_NO_USER_RESOURCE, CLASS_NAME, GetErrMsg(SEC_NO_USER_RESOURCE)
    End If
End Function

'<< sec stub only
Friend Function Initialize(colResources As Collection)

'Accepts a collection of resources as defined in the security_stub.txt file
Dim i As Integer
Dim objKeyValuePair As KeyValuePair

For i = 1 To colResources.Count
    Set objKeyValuePair = colResources.Item(i)
    AddResource objKeyValuePair.Key, vbNullString, vbNullString, objKeyValuePair.Value, Nothing
Next

End Function
