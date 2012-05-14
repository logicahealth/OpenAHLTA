VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "I_IDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<SCR 24303 CC - new class>

Option Explicit

Public Property Let After(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.After = 5
End Property

Public Property Get After() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.After
End Property

Public Property Let Before(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Before = 5
End Property


Public Property Get Before() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Before
End Property

Public Property Let name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
End Property


Public Property Get name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
End Property

Public Property Let ApplicationNCID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ApplicationNCID = 5
End Property

Public Property Get ApplicationNCID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ApplicationNCID
End Property

