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
Attribute VB_Name = "GEMS_Generic_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' WRogers 8/2001
' Declared Value_Text as the Default property for this class
'
'
Private mvarUnit_Number As Long
Private mvarParent      As Long
Private mvarValue_NCID  As Long
Private mvarValue_RSID  As Long
Private mvarValue_Text  As String
Private mvarError_Text  As String
Private mvarType_NCID   As Long
Private mvarOrdinal     As Long
Private mvarUpdateFlag As String

Public Property Get UpdateFlag() As String:           UpdateFlag = mvarUpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mvarUpdateFlag = RHV:         End Property

Public Property Get Unit_Number() As Long:           Unit_Number = mvarUnit_Number: End Property
Public Property Let Unit_Number(ByVal RHV As Long):  mvarUnit_Number = RHV:         End Property

Public Property Get Parent() As Long:                Parent = mvarParent:           End Property
Public Property Let Parent(ByVal RHV As Long):       mvarParent = RHV:              End Property

Public Property Get Value_NCID() As Long:            Value_NCID = mvarValue_NCID:   End Property
Public Property Let Value_NCID(ByVal RHV As Long):   mvarValue_NCID = RHV:          End Property

Public Property Get Value_RSID() As Long:            Value_RSID = mvarValue_RSID:   End Property
Public Property Let Value_RSID(ByVal RHV As Long):   mvarValue_RSID = RHV:          End Property

Public Property Get Value_Text() As String:          Value_Text = mvarValue_Text:   End Property
Attribute Value_Text.VB_UserMemId = 0
Public Property Let Value_Text(ByVal RHV As String): mvarValue_Text = RHV:          End Property

Public Property Get Error_Text() As String:          Error_Text = mvarError_Text:   End Property
Public Property Let Error_Text(ByVal RHV As String): mvarError_Text = RHV:          End Property

Public Property Get Type_NCID() As Long:             Type_NCID = mvarType_NCID:     End Property
Public Property Let Type_NCID(ByVal RHV As Long):    mvarType_NCID = RHV:           End Property

Public Property Get Ordinal() As Long:             Ordinal = mvarOrdinal:     End Property
Public Property Let Ordinal(ByVal RHV As Long):    mvarOrdinal = RHV:           End Property

