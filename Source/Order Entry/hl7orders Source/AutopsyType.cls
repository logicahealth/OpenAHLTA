VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutopsyType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarAutopsyTypeIen As String 'local copy
Private mvarAutopsyTypeName As String 'local copy
Private mvarComments As String 'local copy
Public Property Let Comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comments = 5
    mvarComments = vData
End Property


Public Property Get Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comments
    Comments = mvarComments
End Property



Public Property Let AutopsyTypeName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutopsyTypeName = 5
    mvarAutopsyTypeName = vData
End Property


Public Property Get AutopsyTypeName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutopsyTypeName
    AutopsyTypeName = mvarAutopsyTypeName
End Property



Public Property Let AutopsyTypeIen(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutopsyTypeIen = 5
    mvarAutopsyTypeIen = vData
End Property


Public Property Get AutopsyTypeIen() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutopsyTypeIen
    AutopsyTypeIen = mvarAutopsyTypeIen
End Property



