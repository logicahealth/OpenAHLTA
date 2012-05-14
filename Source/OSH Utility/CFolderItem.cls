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
Attribute VB_Name = "CFolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public AppNCID        As String
Public Ordinal        As Long
Public Image          As String
Public Key            As String
Public psAppName      As String
Public Launch         As String
Public NodeText       As String
Public Tag            As String
Public Index          As Integer
Public ParentEntry    As Long
Public Section        As String
Public Property Get psCaption() As String
   Dim sResult As String
   '''<DEV
   sResult = Trim$(NodeText)
   If sResult = "" Then
      sResult = Trim$(psAppName)
      If sResult = "" Then
         sResult = Trim$(AppNCID)
         If sResult = "" Then
            sResult = "Unknown"
         End If
      End If
   End If
   psCaption = sResult
   '''>DEV ask Bill about caption
End Property


