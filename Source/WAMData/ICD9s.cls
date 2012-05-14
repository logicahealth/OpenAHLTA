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
Attribute VB_Name = "ICD9s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Required. At least 1 ICD9 must be present"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ICD9"
Attribute VB_Ext_KEY = "Member0" ,"ICD9"
Option Explicit
Implements IPersist

'local variable to hold collection
Private mCol As Collection
Private mvarValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
End Sub
Public Function Clear()
    Set mCol = New Collection
End Function

Public Function Add(ByVal ICD_CODE As String, ByVal ICD_DIAGNOSIS_PRIORITY As String) As ICD9
        
    'create a new object
    Dim objNewMember As ICD9
    Dim objICD9 As ICD9
    
    For Each objICD9 In mCol
        If objICD9.ICD_CODE = ICD_CODE Then
            mvarValidator.ValidateDuplicateCode "ICD9s", "ICD_CODE", ICD_CODE
        End If
    Next
    
    Set objNewMember = New ICD9
    objNewMember.Init mvarValidator

    'set the properties passed into the method
    objNewMember.ICD_CODE = ICD_CODE
    objNewMember.ICD_DIAGNOSIS_PRIORITY = ICD_DIAGNOSIS_PRIORITY
    
    mCol.Add objNewMember


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function
'''
'''Public Function AddEx(ByVal ICD_CODE As String, ByVal ICD_DIAGNOSIS_PRIORITY As String, ByVal ACUITY As String, _
'''        ByVal DATE_ONSET As Date, ByVal STATUS As String, ByVal NOTE As String, ByVal NEW_ICD As Boolean) As ICD9
'''
'''    'create a new object
'''    Dim objNewMember As ICD9
'''    Dim objICD9 As ICD9
'''
'''    For Each objICD9 In mCol
'''        If objICD9.ICD_CODE = ICD_CODE Then
'''            mvarValidator.ValidateDuplicateCode "ICD9s", "ICD_CODE", ICD_CODE
'''        End If
'''    Next
'''
'''    Set objNewMember = New ICD9
'''    objNewMember.Init mvarValidator
'''
'''    'set the properties passed into the method
'''    objNewMember.ICD_CODE = ICD_CODE
'''    objNewMember.ICD_DIAGNOSIS_PRIORITY = ICD_DIAGNOSIS_PRIORITY
'''    objNewMember.ACUITY = ACUITY
'''    objNewMember.DATE_ONSET = DATE_ONSET
'''    objNewMember.STATUS = STATUS
'''    objNewMember.NOTE = NOTE
'''    objNewMember.NEW_ICD = NEW_ICD
'''
'''    mCol.Add objNewMember
'''
'''
'''    'return the object created
'''    Set AddEx = objNewMember
'''    Set objNewMember = Nothing
'''
'''
'''End Function

Public Property Get Item(vntIndexKey As Variant) As ICD9
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mvarValidator = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Persistence Block
''''''''''''''''''''''''''''''''''''''''''''''
Private Function IPersist_ChildFactory() As Object
    Dim obj As ICD9
    
    Set obj = New ICD9
    obj.Init mvarValidator
    Set IPersist_ChildFactory = obj
    
End Function
Private Function IPersist_Persist() As String
    IPersist_Persist = PersistCollection(Me, mCol)
End Function
Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = RestoreCollection(Me, mCol, Persist)
End Function

