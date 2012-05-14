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
Attribute VB_Name = "SnoNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:55:29

Private mlMedcinid       As Long
Private msCode           As String
Private miLevel          As Integer
Private msPrefix         As String
Private msNodeKey        As String
Private msDescription    As String
Private miTermtype       As Integer
Private miFlags          As Integer
Private msSubs           As String
Private mbCalculatedDx   As Boolean
Private mobjsnoNodes  As Collection

Public Property Set snoNodes(Value As Collection)

   If mobjsnoNodes Is Nothing Then
      Set mobjsnoNodes = New Collection
   End If

   Set mobjsnoNodes = Value

End Property

Public Property Get snoNodes() As Collection

   If mobjsnoNodes Is Nothing Then
      Set mobjsnoNodes = New Collection
   End If

   Set snoNodes = mobjsnoNodes

End Property

Public Property Get MedcinId() As Long

   MedcinId = mlMedcinid

End Property

Public Property Let MedcinId(Value As Long)

   mlMedcinid = Value

End Property

Public Property Get treeLevel() As Integer

   treeLevel = miLevel

End Property

Public Property Let treeLevel(Value As Integer)

   miLevel = Value

End Property

Public Property Get Prefix() As String

   Prefix = msPrefix

End Property

Public Property Let Prefix(Value As String)

   msPrefix = Value

End Property

Public Property Get code() As String

   code = msCode

End Property

Public Property Let code(Value As String)

   msCode = Value

End Property

Public Property Get Nodekey() As String

   Nodekey = msNodeKey

End Property

Public Property Let Nodekey(Value As String)

   msNodeKey = Value

End Property

Public Property Get Description() As String

   Description = msDescription

End Property

Public Property Let Description(Value As String)

   msDescription = Value

End Property

Public Property Get termtype() As Integer

   termtype = miTermtype

End Property

Public Property Let termtype(Value As Integer)

   miTermtype = Value

End Property

Public Property Get flags() As Integer

   flags = miFlags

End Property

Public Property Let flags(Value As Integer)

   miFlags = Value

End Property

Public Property Get Subs() As String

   Subs = msSubs

End Property

Public Property Let Subs(Value As String)

   msSubs = Value

End Property

Public Property Get calculatedDx() As Boolean

   calculatedDx = mbCalculatedDx

End Property

Public Property Let calculatedDx(Value As Boolean)

   mbCalculatedDx = Value

End Property

