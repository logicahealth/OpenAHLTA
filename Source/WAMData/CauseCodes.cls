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
Attribute VB_Name = "CauseCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist

'local variable to hold collection
Private mcCol As Collection

Private moValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set moValidator = Validatator
End Sub

Public Function Clear()
    Set mcCol = New Collection
End Function

Public Function Add(ByVal sCause_Code As String, ByVal sCause_Code_Description As String) As CauseCode
      
    'create a new object
    Dim oNewMember As CauseCode
    Dim oCauseCode As CauseCode
    
    'Validate duplicates
    For Each oCauseCode In mcCol
        If oCauseCode.Cause_Code = sCause_Code Then
            moValidator.ValidateDuplicateCode "CauseCodes", "Accident Cause Code", sCause_Code
        End If
    Next
    
    moValidator.ValidateColByMaxCount "CauseCodes", "Add", mcCol, 3, "Accident Cause Codes"
    
    Set oNewMember = New CauseCode
    oNewMember.Init moValidator

    'set the properties passed into the method
    oNewMember.Cause_Code = sCause_Code
    oNewMember.Cause_Code_Description = sCause_Code_Description
   
    mcCol.Add oNewMember

    'return the object created
    Set Add = oNewMember
    Set oNewMember = Nothing
    Set oCauseCode = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CauseCode
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mcCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mcCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mcCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mcCol.[_NewEnum]
End Property

''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mcCol = New Collection
End Sub
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mcCol = Nothing
    Set moValidator = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Persistence Block
''''''''''''''''''''''''''''''''''''''''''''''
Private Function IPersist_ChildFactory() As Object
    Dim oCauseCode As CauseCode
    
    Set oCauseCode = New CauseCode
    oCauseCode.Init moValidator
    Set IPersist_ChildFactory = oCauseCode
    
End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = PersistCollection(Me, mcCol)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = RestoreCollection(Me, mcCol, Persist)
End Function




