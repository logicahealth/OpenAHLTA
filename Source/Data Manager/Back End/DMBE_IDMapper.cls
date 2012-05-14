VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMBE_IDMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarApplicationNCID As Long 'local copy
Private mvarFieldName As String 'local copy
Private mvarBeforeValue As String 'local copy
Private mvarAfterValue As String 'local copy

Implements IDMBE.I_IDMapper ''<SCR 24303 CC>

Public Property Let I_IDMapper_After(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.After = 5
    mvarAfterValue = vData
End Property


Public Property Get I_IDMapper_After() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.After
    I_IDMapper_After = mvarAfterValue
End Property



Public Property Let I_IDMapper_Before(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Before = 5
    mvarBeforeValue = vData
End Property


Public Property Get I_IDMapper_Before() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Before
    I_IDMapper_Before = mvarBeforeValue
End Property



Public Property Let I_IDMapper_Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarFieldName = vData
End Property


Public Property Get I_IDMapper_Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    I_IDMapper_Name = mvarFieldName
End Property



Public Property Let I_IDMapper_ApplicationNCID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ApplicationNCID = 5
    mvarApplicationNCID = vData
End Property


Public Property Get I_IDMapper_ApplicationNCID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ApplicationNCID
    I_IDMapper_ApplicationNCID = mvarApplicationNCID
End Property


'****************************************************************
'Delete from down here once everyone implemented the interface.
'****************************************************************




'Public Property Let After(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.After = 5
'    mvarAfterValue = vData
'End Property
'
'
'Public Property Get After() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.After
'    After = mvarAfterValue
'End Property
'
'
'
'Public Property Let Before(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Before = 5
'    mvarBeforeValue = vData
'End Property
'
'
'Public Property Get Before() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.Before
'    Before = mvarBeforeValue
'End Property
'
'
'
'Public Property Let Name(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Name = 5
'    mvarFieldName = vData
'End Property
'
'
'Public Property Get Name() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.Name
'    Name = mvarFieldName
'End Property
'
'
'
'Public Property Let ApplicationNCID(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.ApplicationNCID = 5
'    mvarApplicationNCID = vData
'End Property
'
'
'Public Property Get ApplicationNCID() As Long
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.ApplicationNCID
'    ApplicationNCID = mvarApplicationNCID
'End Property
'
