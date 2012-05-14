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
Attribute VB_Name = "Drug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=Drug
'--- 4/1/2004 2:48:20 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'
Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Drug"


Private mContextNCID As String  '
Private mDrugType As DrugType   '
Private mDose As DrugDose   '
Private mNationalDrugCode As String '
Private mSensitive As Boolean   '
Private mTextComments As String '
Private mCodedComments As String    '
Private mBaseDrug As Boolean    '

Private Sub Class_Initialize()
  Set mDrugType = New DrugType
  Set mDose = New DrugDose
End Sub

Public Property Get ContextNCID() As String
    ContextNCID = mContextNCID
End Property
Public Property Let ContextNCID(ByVal value As String)
    mContextNCID = value
End Property
Public Property Get DrugType() As DrugType
    Set DrugType = mDrugType
End Property
Public Property Set DrugType(ByRef value As DrugType)
    Set mDrugType = value
End Property
Public Property Get Dose() As DrugDose
    Set Dose = mDose
End Property
Public Property Set Dose(ByRef value As DrugDose)
    Set mDose = value
End Property
Public Property Get NationalDrugCode() As String
    NationalDrugCode = mNationalDrugCode
End Property
Public Property Let NationalDrugCode(ByVal value As String)
    mNationalDrugCode = value
End Property
Public Property Get Sensitive() As Boolean
    Sensitive = mSensitive
End Property
Public Property Let Sensitive(ByVal value As Boolean)
    mSensitive = value
End Property
Public Property Get TextComments() As String
    TextComments = mTextComments
End Property
Public Property Let TextComments(ByVal value As String)
    mTextComments = value
End Property
Public Property Get CodedComments() As String
    CodedComments = mCodedComments
End Property
Public Property Let CodedComments(ByVal value As String)
    mCodedComments = value
End Property
Public Property Get BaseDrug() As Boolean
    BaseDrug = mBaseDrug
End Property
Public Property Let BaseDrug(ByVal value As Boolean)
    mBaseDrug = value
End Property

Private Function iObjectPersist_Copy() As Object

On Error GoTo ErrHandler

    Dim NewCopy As Drug
    Dim Persist As iObjectPersist
    
    Set NewCopy = New Drug
    
    With NewCopy
        .ContextNCID = mContextNCID
        Set Persist = mDrugType
        Set .DrugType = Persist.Copy
        Set Persist = mDose
        Set .Dose = Persist.Copy
        .NationalDrugCode = mNationalDrugCode
        .Sensitive = mSensitive
        .TextComments = mTextComments
        .CodedComments = mCodedComments
        .BaseDrug = mBaseDrug
    End With

    Set iObjectPersist_Copy = NewCopy

Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsDrug.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)

On Error GoTo ErrHandler

    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ContextNCID", mContextNCID
        .WriteObject "DrugType", mDrugType
        .WriteObject "Dose", mDose
        .WriteProperty "NationalDrugCode", mNationalDrugCode
        .WriteProperty "Sensitive", mSensitive
        .WriteProperty "TextComments", mTextComments
        .WriteProperty "CodedComments", mCodedComments
        .WriteProperty "BaseDrug", mBaseDrug
    End With
Exit Sub
ErrHandler:
    Err.Raise Err.Number, "DTOsDrug.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mContextNCID = .ReadProperty("ContextNCID")
        .ReadObject "DrugType", DrugType
        .ReadObject "Dose", Dose
        mNationalDrugCode = .ReadProperty("NationalDrugCode")
        mSensitive = .ReadProperty("Sensitive")
        mTextComments = .ReadProperty("TextComments")
        mCodedComments = .ReadProperty("CodedComments")
        mBaseDrug = .ReadProperty("BaseDrug")
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

