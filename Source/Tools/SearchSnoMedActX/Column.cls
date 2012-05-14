VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Column for Display on Selection Tab"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit


Private mvarKey As String 'local copy
Private mvarCaption As String 'local copy
Private mvarSorted As Boolean 'local copy
Private mvarSort As enumListSortOrderConstants 'local copy
Private mvarVisible As Boolean 'local copy
Private mvarCharactersWide As Long

Public Property Let Visible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Visible = 5
    mvarVisible = vData
End Property


Public Property Get Visible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
End Property

Public Property Let Sort(ByVal vData As enumListSortOrderConstants)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sort = 5
    mvarSort = vData
End Property


Public Property Get Sort() As enumListSortOrderConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sort
    Sort = mvarSort
End Property


Public Property Let Sorted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sorted = 5
    mvarSorted = vData
End Property


Public Property Get Sorted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sorted
    Sorted = mvarSorted
End Property



Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Key = mvarKey
End Property

Public Property Let CharactersWide(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarCharactersWide = vData
End Property


Public Property Get CharactersWide() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    CharactersWide = mvarCharactersWide
End Property


