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
Attribute VB_Name = "colGCDxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsGCDx"
Attribute VB_Ext_KEY = "Member0" ,"clsGCDx"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(GCDx As clsGCDx, Optional sKey As String) As clsGCDx

'This sorts the collection in the following order
'By drc ascending, by datetime ascending
'Example: For dxs, display the oldest first, therefore put it at the bottom of the collection.
'The items at the bottom display first!
'GCDx1: datetime-1/4/03 10:30:23am, drc-1
'GCDx2: datetime-1/4/03 10:32:10am, drc-2
'GCDx3: datetime-1/2/03 whocares, drc 3

'Dim i As Integer
'Dim j As Integer
'Dim bFoundHigherDRC As Boolean
Dim iCounter As Integer
Dim bAdded As Boolean
    
    bAdded = False
    
    For iCounter = mCol.Count To 1 Step -1
        If GCDx.DRC > mCol(iCounter).DRC Then
            'The new gcDx has a higher drc so add it after the current one
            mCol.Add GCDx, , , iCounter
            bAdded = True
            Exit For
        ElseIf GCDx.DRC = mCol(iCounter).DRC And GCDx.ModifiedDateTime < mCol(iCounter).ModifiedDateTime Then
            'The new gcDx has the same drc and an older date so add it after the current one.
            mCol.Add GCDx, , , iCounter
            bAdded = True
            Exit For
        End If
    Next
    
    If bAdded = False Then
        If mCol.Count = 0 Then
            mCol.Add GCDx
        Else
            'Add it to the top of the collection
            mCol.Add GCDx, , 1
        End If
    End If
    
'


'    'Add the items in order by datetime decending
'    For i = 1 To mCol.Count
'        If mCol(i).ModifiedDateTime <= GCDx.ModifiedDateTime Then
'            Exit For
'        End If
'    Next i
'
'    'Go through the items backwards (oldest first and find a lower drc in the list).
'    For j = mCol.Count To 1 Step -1
'        If mCol(j).DRC < GCDx.DRC Then
'            bFoundHigherDRC = True
'            Exit For
'        Else
'            bFoundHigherDRC = False
'        End If
'    Next j
'
'    If bFoundHigherDRC = True Then
'        'Add the higher drc after the position of the drc found above
'        mCol.Add GCDx, , , j
'
'    ElseIf i > mCol.Count Then
'        'Add the GCDx to the end of the collection
'        mCol.Add GCDx
'    Else
'        'Add the new GCDx before the GCDx at position i
'        mCol.Add GCDx, , i
'    End If
    
    Set Add = GCDx


End Function

Public Property Get Item(vntIndexKey As Variant) As clsGCDx
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

