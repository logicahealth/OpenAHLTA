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
Attribute VB_Name = "ListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ListItem"
Attribute VB_Ext_KEY = "Member0" ,"ListItem"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mSelectTab As SelectTab

Friend Sub Init(NewSelectTab As SelectTab)
    Set mSelectTab = NewSelectTab
End Sub
Public Property Get SelectTab() As SelectTab
    Set SelectTab = mSelectTab
End Property
Friend Function AddListItem(ByRef NewListItem As ListItem) As ListItem
    'create a new object
    Dim i As Long
    
    On Error GoTo errXIT:

    mCol.Add NewListItem

    'return the object created
    Set AddListItem = NewListItem

    Exit Function
errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Friend Function AddArray(ByRef Columns As Variant, Values As Variant) As ListItem
    'create a new object
    Dim objNewMember As ListItem
    Dim i As Long
    
    On Error GoTo errXIT:
    
    Set objNewMember = New ListItem
    
    If UBound(Columns) <> UBound(Values) Then
        Err.Raise vbObjectError, App.Title, "Values, Columns array mismatch."
    End If
    
    For i = LBound(Columns) To UBound(Columns)
        objNewMember.Items.Add CStr(Values(i)), CStr(Columns(i))
    Next

    mCol.Add objNewMember

    'return the object created
    Set AddArray = objNewMember
    Set objNewMember = Nothing

    Exit Function
errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description
    

End Function
Public Function Add(ParamArray Values()) As ListItem
    'create a new object
    Dim objNewMember As ListItem
    Dim i As Long
    Dim c As Column
    Dim sz As String
    
    On Error GoTo errXIT:
    
    Set objNewMember = New ListItem
    
    If UBound(Values) > mSelectTab.Columns.Count - 1 Then
        Err.Raise vbObjectError, App.Title, "Attemted to add more Items than present in Columns collection."
    End If
    
    For i = 1 To mSelectTab.Columns.Count
        sz = ""
        If i - 1 <= UBound(Values) Then
            sz = CStr(Values(i - 1))
        End If
    
        objNewMember.Items.Add sz, mSelectTab.Columns(i).Key
    Next

    mCol.Add objNewMember

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    Exit Function
errXIT:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Property Get Item(vntIndexKey As Variant) As ListItem
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

