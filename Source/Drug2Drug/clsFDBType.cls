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
Attribute VB_Name = "clsFDBType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum AllergyTypeEnum
    AllergenGroup = 1
    DispensableDrug = 4
    Ingredient = 6
    Compound = 16
    DrugName = 2
    ManufacturedDrug = 14
    PackagedDrug = 5
    RoutedDrug = 3
End Enum

'local variable(s) to hold property value(s)
Private mvarAllergyType As AllergyTypeEnum 'local copy
Private mvarFDBTable As String 'local copy
Private mvarPK As String 'local copy
Private mvarDescriptionColumn As String 'local copy
Public Property Let DescriptionColumn(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DescriptionColumn = 5
    mvarDescriptionColumn = vData
End Property


Public Property Get DescriptionColumn() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DescriptionColumn
    DescriptionColumn = mvarDescriptionColumn
End Property



Public Property Let PK(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PK = 5
    mvarPK = vData
End Property


Public Property Get PK() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PK
    PK = mvarPK
End Property



Public Property Let FDBTable(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FDBTable = 5
    mvarFDBTable = vData
End Property


Public Property Get FDBTable() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FDBTable
    FDBTable = mvarFDBTable
End Property



Public Property Let AllergyType(ByVal vData As AllergyTypeEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AllergyType = 5
    mvarAllergyType = vData
End Property


Public Property Get AllergyType() As AllergyTypeEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllergyType
    AllergyType = mvarAllergyType
End Property



