'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHCSII_Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mParams As Collection
Private mStoredProcedureName As String
Public Property Get SPName() As String
    SPName = mStoredProcedureName
End Property
Public Property Let SPName(ByVal vData As String)
    mStoredProcedureName = vData
End Property

Public Function AddParam(ByVal Is_Cursor As Long, _
                            ByVal Is_Output As Long, _
                            ByVal Info_Type As String, _
                            ByVal Param_Name As String) As Boolean

    Dim newParam As CHCSII_SPparam
    Set newParam = New CHCSII_SPparam
    
    newParam.IsCursor = Is_Cursor
    newParam.IsOutput = Is_Output
    newParam.InfoType = Info_Type
    newParam.Param = Param_Name
    
    If Is_Output = False And Param_Name = "" Then
        newParam.Param = "NULL"
    End If
    
    mParams.Add newParam
    AddParam = True

End Function
Public Property Get Item(vntIndexKey As Variant) As CHCSII_SPparam

  If vntIndexKey = 0 Then Exit Property
  Set Item = mParams(CInt(vntIndexKey))
  
End Property

Public Property Get Count() As Long
    Count = mParams.Count
End Property

Private Sub Class_Initialize()
    Set mParams = New Collection
End Sub

Private Sub Class_Terminate()
    Set mParams = Nothing
End Sub

