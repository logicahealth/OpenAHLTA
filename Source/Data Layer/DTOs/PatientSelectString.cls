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
Attribute VB_Name = "PatientSelectString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=PatientSelectString
'--- 2/17/2004 6:21:50 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "PatientSelectString"


Private mID As String   '
Private mValue As String    '

Public Enum enumSelectStringCategory
  osPatientCategory = 0
  osFMP = 1
  osRank = 2
  osMaritalStatus = 3
  osRace = 4
  osEnrollmentFacility = 5
  osSpecialtyCode = 6
  osReligion = 7
  osService = 8
  osPhoneType = 9
  osAddressType = 10
  osGender = 11
  osUnknownCategory = 12
  osRecordsMaintainedAt = 13
  osACVEnrollmentStatus = 14
End Enum


Public Property Get ID() As String
    ID = mID
End Property
Public Property Let ID(ByVal value As String)
    mID = value
End Property
Public Property Get value() As String
    value = mValue
End Property
Public Property Let value(ByVal value As String)
    mValue = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As PatientSelectString
    On Error GoTo ErrHandler
    
    Set NewCopy = New PatientSelectString
    
    With NewCopy
        .ID = mID
        .value = mValue
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsPatientSelectString.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)
    
    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ID", mID
        .WriteProperty "Value", mValue
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsPatientSelectString.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mID = .ReadProperty("ID")
        mValue = .ReadProperty("Value")
    End With

    iObjectPersist_Restore = Not Failed
  
    Exit Function
  
ErrHandler:
    Failed = True
    Resume Next

End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder

