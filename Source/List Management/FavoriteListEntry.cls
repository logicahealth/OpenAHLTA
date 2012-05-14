VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FavoriteListEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarNCID As String 'local copy
'local variable(s) to hold property value(s)
Private mvarICD9Code As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCPT4Code As String 'local copy
Public Property Let CPT4Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPT4Code = 5
    mvarCPT4Code = vData
End Property


Public Property Get CPT4Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPT4Code
    CPT4Code = mvarCPT4Code
End Property



Public Property Let ICD9Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICD9Code = 5
    mvarICD9Code = vData
End Property


Public Property Get ICD9Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICD9Code
    ICD9Code = mvarICD9Code
End Property



Public Property Let NCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NCID = 5
    mvarNCID = vData
End Property


Public Property Get NCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NCID
    NCID = mvarNCID
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



