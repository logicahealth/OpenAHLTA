VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Res_Atrrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_lResourceID As Long
Private m_sResourceName As String
Private m_colAttributes As Collection    '  of KeyValuePair items

Friend Property Let ResourceID(ByVal ResourceID As String):    m_lResourceID = ResourceID: End Property
Public Property Get ResourceID() As String:    ResourceID = m_lResourceID: End Property

Friend Property Let ResourceName(ByVal ResourceName As String):    m_sResourceName = ResourceName: End Property
Public Property Get ResourceName() As String:    ResourceName = m_sResourceName: End Property
Friend Property Set Attributes(colAttributes As Collection):    Set m_colAttributes = colAttributes: End Property
Public Property Get Attributes() As Collection:    Set Attributes = m_colAttributes: End Property

Public Function AddAttribute(NewAttrib As KeyValuePair) As Boolean
    Attributes.Add NewAttrib, NewAttrib.Key
End Function

Public Function RemoveAttribute(sAttribKey As String) As Boolean
    Dim oKeyVal As KeyValuePair
    Dim i As Integer
    
    For i = 1 To Attributes.count
        Set oKeyVal = Attributes.Item(i)
        If UCase(oKeyVal.Key) = UCase$(sAttribKey) Then
            Set oKeyVal = Nothing
            Attributes.Remove (i)
            RemoveAttribute = True
            Exit Function
        End If
    Next i
    
    Set oKeyVal = Nothing
    RemoveAttribute = False
End Function

Private Sub Class_Initialize()
    Set m_colAttributes = New Collection
End Sub
