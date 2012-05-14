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
Attribute VB_Name = "clsFrequency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCode        As String
Private mDescribtion As String

Public Property Let code(ByVal vData As String)
    mCode = vData
End Property
Public Property Get code() As String
    code = mCode
End Property
Public Property Let Describtion(ByVal vData As String)
    mDescribtion = vData
End Property
Public Property Get Describtion() As String
    Describtion = mDescribtion
End Property

Public Function GetFrequencys() As Collection

Dim rs      As Recordset
Dim strSQL  As String
Dim colFrqs As Collection
Dim freq    As clsFrequency

On Error GoTo ErrHnd

    Set colFrqs = New Collection
    
    strSQL = "select * from Frequency_code "
    
    Set rs = goBB.I_DAS.OpenRecordset(strSQL)

    Do While Not rs.EOF
        Set freq = New clsFrequency
        freq.code = rs!FREQUENCY_CODE & vbNullString
        freq.Describtion = rs!FREQUENCY_DESC & vbNullString
        colFrqs.Add freq
        Set freq = Nothing
        rs.MoveNext
    Loop
    
    rs.Close
    Set GetFrequencys = colFrqs
    Exit Function
    
ErrHnd:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsDispense.retrieveData", "clsDispense", vbExclamation
End Function
