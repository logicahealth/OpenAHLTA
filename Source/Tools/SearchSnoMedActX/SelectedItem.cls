VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectedItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Container for Prcedure Name, NCID, and MedcinID returned to developer when using frmSelect"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDescription As String   'local copy
Private mvarNCID As String          'local copy
Private mvarMedcinID As String      'local copy
'local variable(s) to hold property value(s)
Private mvarICD9 As String 'local copy
Private mvarCPT4 As String 'local copy
'local variable(s) to hold property value(s)
Private mvarPrefix As String 'local copy
Friend Property Let Prefix(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prefix = 5
    mvarPrefix = vData
End Property


Public Property Get Prefix() As String
Attribute Prefix.VB_Description = "Medcin Prefix used to determine: \r\nH=> History Of\r\nF=> Family History of"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prefix
    Prefix = mvarPrefix
End Property



Friend Property Let CPT4(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPT4 = 5
    mvarCPT4 = vData
End Property


Public Property Get CPT4() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPT4
    CPT4 = mvarCPT4
End Property



Friend Property Let ICD9(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICD9 = 5
    mvarICD9 = vData
End Property


Public Property Get ICD9() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICD9
    ICD9 = mvarICD9
End Property



Friend Property Let medcinid(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MedcinID = 5
    mvarMedcinID = vData
End Property


Public Property Get medcinid() As String
Attribute medcinid.VB_Description = "The selected MedcinID"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MedcinID
    medcinid = mvarMedcinID
End Property



Friend Property Let NCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NCID = 5
    mvarNCID = vData
End Property


Public Property Get NCID() As String
Attribute NCID.VB_Description = "The selected NCID code."
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NCID
    NCID = mvarNCID
End Property



Friend Property Let description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get description() As String
Attribute description.VB_Description = "The Selected Procedure name"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    description = mvarDescription
End Property



