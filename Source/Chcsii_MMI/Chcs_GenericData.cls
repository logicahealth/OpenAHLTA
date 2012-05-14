VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Chcs_GenericData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolGenData As Collection
Private mbBOF       As Boolean
Private mbEOF       As Boolean
Private mlIndex     As Long

Implements IGenericData

Private Sub Class_Initialize()
        
    Set mcolGenData = New Collection
    mbBOF = True
    mbEOF = True
        
End Sub

Private Sub Class_Terminate()

    Set mcolGenData = Nothing

End Sub

Private Property Get IGenericData_BOF() As Boolean

    If mlIndex = 0 Then IGenericData_BOF = True

End Property

Private Property Let IGenericData_Container(ByVal RHS As Variant)

End Property

Private Property Get IGenericData_Container() As Variant

End Property

Private Property Get IGenericData_EOF() As Boolean

    If mlIndex > mcolGenData.Count Or mcolGenData.Count = 0 Then IGenericData_EOF = True

End Property

Private Property Get IGenericData_ErrorText() As String

End Property

Private Property Let IGenericData_ErrorText(ByVal RHS As String)

End Property

Private Sub IGenericData_GotoFirst()

    mlIndex = 1
        
End Sub

Private Sub IGenericData_GotoLast()

    mlIndex = mcolGenData.Count

End Sub

Private Sub IGenericData_GotoNext()

    If mlIndex <= mcolGenData.Count Then
        mlIndex = mlIndex + 1
        mbBOF = False
    End If
            
    If mlIndex > mcolGenData.Count Then mbEOF = True

End Sub

Private Sub IGenericData_GotoPrevious()

    If mlIndex >= 1 Then
        mlIndex = mlIndex - 1
        mbEOF = False
    End If
    
    If mlIndex = 0 Then mbBOF = True

End Sub

Private Property Let IGenericData_StoreTime(ByVal RHS As Date)

End Property

Private Property Get IGenericData_StoreTime() As Date

End Property

Private Property Let IGenericData_Type(ByVal RHS As String)
    Dim oGenData As New GEMS_Generic_Data

    If mbEOF Then
        mlIndex = mcolGenData.Count + 1
        oGenData.Type_NCID = RHS
        
        mcolGenData.Add oGenData, CStr(mlIndex)
        mbBOF = False
        mbEOF = False
    End If

End Property

Private Property Get IGenericData_Type() As String
    
    IGenericData_Type = mcolGenData.Item(CStr(mlIndex)).Type_NCID

End Property

Private Property Let IGenericData_ValueNcid(ByVal RHS As String)
    mcolGenData.Item(CStr(mlIndex)).Value_NCID = RHS
End Property

Private Property Get IGenericData_ValueNcid() As String

    IGenericData_ValueNcid = mcolGenData.Item(CStr(mlIndex)).Value_NCID

End Property

Private Property Let IGenericData_ValueRsid(ByVal RHS As String)

End Property

Private Property Get IGenericData_ValueRsid() As String

End Property

Private Property Let IGenericData_ValueText(ByVal RHS As String)
    
    mcolGenData.Item(CStr(mlIndex)).Value_Text = RHS

End Property

Private Property Get IGenericData_ValueText() As String

    IGenericData_ValueText = mcolGenData.Item(CStr(mlIndex)).Value_Text

End Property
