VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Allergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Allergies"
Option Explicit

Private mvarsFMP As String 'local copy
Private mvarsSSN As String 'local copy
Private mvarsAllergenNCID As String 'local copy
Private mvarsSeverityNCID As String 'local copy
'local variable(s) to hold property value(s)
Private mvarIEN As Long 'local copy
Public Property Let lIEN(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IEN = 5
    mvarIEN = vData
End Property


Public Property Get lIEN() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IEN
    lIEN = mvarIEN
End Property







Public Property Let sSeverity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSeverityNCID = 5
    mvarsSeverityNCID = vData
End Property


Public Property Get sSeverity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSeverityNCID
    sSeverity = mvarsSeverityNCID
End Property


Public Property Let sAllergen(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAllergenNCID = 5
    mvarsAllergenNCID = vData
End Property


Public Property Get sAllergen() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAllergenNCID
    sAllergen = mvarsAllergenNCID
End Property



Public Property Let sSSN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSSN = 5
    mvarsSSN = vData
End Property


Public Property Get sSSN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSSN
    sSSN = mvarsSSN
End Property



Public Property Let sFMP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFMP = 5
    mvarsFMP = vData
End Property


Public Property Get sFMP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFMP
    sFMP = mvarsFMP
End Property



