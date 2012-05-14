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
Attribute VB_Name = "colPatTX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Items As Collection
' Add an item at its correct position.
Public Function Add(ByVal new_item As patTx) As patTx
Dim i As Integer

    ' See where the item belongs.
    For i = 1 To m_Items.Count
        If m_Items(i).dModifiedDateTime >= new_item.dModifiedDateTime Then Exit For
    Next i

    ' Insert the item.
    If i > m_Items.Count Then
        ' Add at the end.
        m_Items.Add new_item
    Else
        ' Add at the right position.
        m_Items.Add new_item, , i
    End If
    
    Set Add = new_item
    
End Function
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = m_Items.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    m_Items.Remove vntIndexKey
End Sub
Public Property Get Item(vntIndexKey As Variant) As patTx
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = m_Items(vntIndexKey)
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_Items.[_NewEnum]
End Property

' Create the new collection.
Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub


