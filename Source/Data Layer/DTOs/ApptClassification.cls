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
Attribute VB_Name = "ApptClassification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=ApptClassification
'--- 3/18/2004 11:28:56 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "ApptClassification"


Private mClassification As String   '
Private mMeprsCode As String    '
Private mInpatientIEN As String '
Private mDescription As String  '
Private mRegisterID As String   '
Private mObservation As String  '



Public Property Get Classification() As String
    Classification = mClassification
End Property
Public Property Let Classification(ByVal value As String)
    mClassification = value
End Property
Public Property Get MeprsCode() As String
    MeprsCode = mMeprsCode
End Property
Public Property Let MeprsCode(ByVal value As String)
    mMeprsCode = value
End Property
Public Property Get InpatientIEN() As String
    InpatientIEN = mInpatientIEN
End Property
Public Property Let InpatientIEN(ByVal value As String)
    mInpatientIEN = value
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(ByVal value As String)
    mDescription = value
End Property
Public Property Get RegisterID() As String
    RegisterID = mRegisterID
End Property
Public Property Let RegisterID(ByVal value As String)
    mRegisterID = value
End Property
Public Property Get Observation() As String
    Observation = mObservation
End Property
Public Property Let Observation(ByVal value As String)
    mObservation = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As ApptClassification
    Set NewCopy = New ApptClassification
    
    With NewCopy
        .Classification = mClassification
        .MeprsCode = mMeprsCode
        .InpatientIEN = mInpatientIEN
        .Description = mDescription
        .RegisterID = mRegisterID
        .Observation = mObservation
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "Classification", mClassification
        .WriteProperty "MeprsCode", mMeprsCode
        .WriteProperty "InpatientIEN", mInpatientIEN
        .WriteProperty "Description", mDescription
        .WriteProperty "RegisterID", mRegisterID
        .WriteProperty "Observation", mObservation
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mClassification = .ReadProperty("Classification")
        mMeprsCode = .ReadProperty("MeprsCode")
        mInpatientIEN = .ReadProperty("InpatientIEN")
        mDescription = .ReadProperty("Description")
        mRegisterID = .ReadProperty("RegisterID")
        mObservation = .ReadProperty("Observation")
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

