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
Attribute VB_Name = "Allergen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=Allergen
'--- 1/21/2004 6:23:27 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Allergen"


Private mName As String '
Private mNCID As String '

'--- end init / terminate procedures

'--- begin property lets/gets
Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(Value As String)
    mName = Value
End Property
Public Property Get NCID() As String
    NCID = mNCID
End Property
Public Property Let NCID(Value As String)
    mNCID = Value
End Property
'--- end property lets/gets




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Allergen
    Set NewCopy = New Allergen
    
    With NewCopy
        .Name = mName
        .NCID = mNCID
    End With

    Set iObjectPersist_Copy = NewCopy

End Function
'--- end copy procedure




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "Name", mName
        .WriteProperty "NCID", mNCID
    End With
End Sub
'--- end persist procedure




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mName = .ReadProperty("Name")
        mNCID = .ReadProperty("NCID")
    End With

    iObjectPersist_Restore = Not Failed
  
    Exit Function
  
ErrHandler:
    Failed = True
    Resume Next

End Function
'--- end restore procedure


Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder

