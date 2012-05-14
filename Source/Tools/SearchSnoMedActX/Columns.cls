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
Attribute VB_Name = "Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Column"
Attribute VB_Ext_KEY = "Member0" ,"Column"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mUserControl As UserControl

Public Enum enumListSortOrderConstants
    SortAscending = ListSortOrderConstants.lvwAscending
    SortDescending = ListSortOrderConstants.lvwDescending
End Enum

Friend Sub Init(NewUserControl As UserControl)
    Set mUserControl = NewUserControl
End Sub

Public Function Add(Caption As String, Optional CharactersWide As Long = 10, Optional Visible As Boolean = True, _
    Optional Sort As enumListSortOrderConstants = lvwAscending, Optional Sorted As Boolean = False) As Column
    'create a new object
    Dim objNewMember As Column
    Set objNewMember = New Column


    'set the properties passed into the method
    objNewMember.Caption = Caption
    objNewMember.Sorted = Sorted
    objNewMember.Sort = Sort
    objNewMember.Visible = Visible
    objNewMember.Key = Caption
    objNewMember.CharactersWide = CharactersWide
    
    mCol.Add objNewMember, Caption


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Column
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


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

