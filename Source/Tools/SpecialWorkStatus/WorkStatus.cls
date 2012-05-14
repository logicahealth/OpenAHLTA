VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum EnumQualified   'values Used In DB
    eqDisqualified = 0
    eqQualified = 1
    eqUnspecified = -1
End Enum
Public Enum EnumSelected    'values Used In DB
    esSelected = -1
    esRemoved = 0
End Enum
Public Enum EnumInsert
    eiInsert = 1
    eiUpate = 2
End Enum


Public NCID As String
Public Selected As EnumSelected
Public Selected_Initial As EnumSelected
Public Qualified As EnumQualified
Public Qualified_Initial As EnumQualified

Public DESCRIPTION As String
Public Insert As EnumInsert



