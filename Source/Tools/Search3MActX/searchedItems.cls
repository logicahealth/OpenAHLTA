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
   MultiUse = -1                       ' True
   Persistable = 0                     ' NotPersistable
   DataBindingBehavior = 0             ' vbNone
   DataSourceBehavior  = 0             ' vbNone
   MTSTransactionMode  = 0             ' NotAnMTSObject
END

Attribute VB_Name = "searchedItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:54:25

Public Key  As String

'local variable(s) to hold property value(s)
Private mvarstrCode              As String    ' local copy
Private mvarstrLongDescription   As String    ' local copy
Private mvarstrShortDescription  As String    ' local copy
Private mvarlngNCID              As Long      ' local copy

Public Property Let strshortDescription(ByVal vData As String)

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strShortDescription = 5
   mvarstrShortDescription = vData

End Property

Public Property Get strshortDescription() As String

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strShortDescription
   strshortDescription = mvarstrShortDescription

End Property

Public Property Let strLongDescription(ByVal vData As String)

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strLongDescription = 5
   mvarstrLongDescription = vData

End Property

Public Property Get strLongDescription() As String

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strLongDescription
   strLongDescription = mvarstrLongDescription

End Property

Public Property Let strCode(ByVal vData As String)

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCode = 5
   mvarstrCode = vData

End Property

Public Property Get strCode() As String

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCode
   strCode = mvarstrCode

End Property

Public Property Let lngNCID(ByVal vData As Long)

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strCode = 5
   mvarlngNCID = vData

End Property

Public Property Get lngNCID() As Long

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strCode
   lngNCID = mvarlngNCID

End Property

