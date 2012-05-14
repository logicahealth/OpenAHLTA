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
Attribute VB_Name = "XferTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Index As Long
Public TableName As String
Public SelectPart As String
Public WherePart As String
Public WherePatientPart As String
Public FromPart As String
Public AdditionalSort As String
Public PatientUniqueFields As Dictionary
Public KeyFields As Dictionary
Public Fields As XferFields



Public Function AddField(strFieldName As String)
    Dim objField As XferField
    
    Set objField = New XferField
    objField.FieldName = strFieldName
    Me.Fields.Add objField
    Set objField = Nothing
    
End Function

Private Sub Class_Initialize()
    Set PatientUniqueFields = New Dictionary
    Set KeyFields = New Dictionary
    Set Fields = New XferFields
End Sub

Private Sub Class_Terminate()
    Set PatientUniqueFields = Nothing
    Set KeyFields = Nothing
    Set Fields = Nothing
End Sub
