VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarFolder_ID As String 'local copy
Private mvarFolder_Name As String 'local copy
Private mvarFolder_Parent_ID As String 'local copy
Private mvarUSER_NCID As String 'local copy
Private mvarFolder_Type As String 'local copy
Private mvarExpanded As Boolean 'local copy
Private mvarLoaded As Boolean 'local copy
Private mvarCustomFolder As Boolean 'SCR-44264 RBELL 12/23/03 Define flag value to determine the origin of folder

Public Property Let Loaded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Loaded = 5
    mvarLoaded = vData
End Property


Public Property Get Loaded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Loaded
    Loaded = mvarLoaded
End Property



Public Property Let Expanded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Expanded = 5
    mvarExpanded = vData
End Property


Public Property Get Expanded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Expanded
    Expanded = mvarExpanded
End Property



Public Property Let Folder_Type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Folder_Type = 5
    mvarFolder_Type = vData
End Property


Public Property Get Folder_Type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Folder_Type
    Folder_Type = mvarFolder_Type
End Property



Public Property Let USER_NCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.USER_NCID = 5
    mvarUSER_NCID = vData
End Property


Public Property Get USER_NCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.USER_NCID
    USER_NCID = mvarUSER_NCID
End Property



Public Property Let Folder_Parent_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Folder_Parent_ID = 5
    mvarFolder_Parent_ID = vData
End Property


Public Property Get Folder_Parent_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Folder_Parent_ID
    Folder_Parent_ID = mvarFolder_Parent_ID
End Property



Public Property Let Folder_Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Folder_Name = 5
    mvarFolder_Name = vData
End Property


Public Property Get Folder_Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Folder_Name
    Folder_Name = mvarFolder_Name
End Property



Public Property Let Folder_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Folder_ID = 5
    mvarFolder_ID = vData
End Property


Public Property Get Folder_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Folder_ID
    Folder_ID = mvarFolder_ID
End Property




Public Property Get CustomFolder() As Boolean
    'SCR-44264 RBELL 12/23/03 Expose flag value to determine the origin of folder
    CustomFolder = mvarCustomFolder
End Property

Public Property Let CustomFolder(ByVal RHV As Boolean)
    'SCR-44264 RBELL 12/23/03 Allow user to set flag value to determine the origin of folder
    mvarCustomFolder = RHV
End Property
