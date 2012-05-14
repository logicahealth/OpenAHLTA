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
Attribute VB_Name = "DrugDose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=DrugDose
'--- 4/1/2004 2:55:49 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'
Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "DrugDose"


Private mDataType As Integer    '
Private mDescription As String  '
Private mUpperLimit As String  '
Private mLowerLimit As String  '
Private mUnit As ConceptValue   '


Private Sub Class_Initialize()
  Set mUnit = New ConceptValue
End Sub

Public Property Get DataType() As Integer
    DataType = mDataType
End Property
Public Property Let DataType(ByVal value As Integer)
    mDataType = value
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(ByVal value As String)
    mDescription = value
End Property
Public Property Get UpperLimit() As String
    UpperLimit = mUpperLimit
End Property
Public Property Let UpperLimit(ByVal value As String)
    mUpperLimit = value
End Property
Public Property Get LowerLimit() As String
    LowerLimit = mLowerLimit
End Property
Public Property Let LowerLimit(ByVal value As String)
    mLowerLimit = value
End Property
Public Property Get Unit() As ConceptValue
    Set Unit = mUnit
End Property
Public Property Set Unit(ByRef value As ConceptValue)
    Set mUnit = value
End Property

Private Function iObjectPersist_Copy() As Object
    
On Error GoTo ErrHandler
    
    Dim NewCopy As DrugDose
    Dim Persist As iObjectPersist
    
    Set NewCopy = New DrugDose
    
    With NewCopy
        .DataType = mDataType
        .Description = mDescription
        .UpperLimit = mUpperLimit
        .LowerLimit = mLowerLimit
        Set Persist = mUnit
        Set .Unit = Persist.Copy
    End With

    Set iObjectPersist_Copy = NewCopy

Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsDrugDose.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Sub iObjectPersist_Persist(Writer As iWriter)

On Error GoTo ErrHandler

    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "DataType", mDataType
        .WriteProperty "Description", mDescription
        .WriteProperty "UpperLimit", mUpperLimit
        .WriteProperty "LowerLimit", mLowerLimit
        .WriteObject "Unit", mUnit
    End With

Exit Sub
ErrHandler:
    Err.Raise Err.Number, "DTOsDrugDose.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mDataType = .ReadProperty("DataType")
        mDescription = .ReadProperty("Description")
        mUpperLimit = .ReadProperty("UpperLimit")
        mLowerLimit = .ReadProperty("LowerLimit")
        .ReadObject "Unit", Unit
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

