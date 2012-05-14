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
Attribute VB_Name = "HistoryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngSnoID As String             ' Current SnoID at change
Private mintViewMode As Integer          ' Current view mode
Private mintEntryTab As Integer         ' Current entry tab
Private mintListsize As Integer         ' Current list size
Private mstrSarg As String               ' search string, key array, or Dx array
Private mcolTemplateItems As Collection           ' Template collection, if applicable
Private menumScreenMode As ModeEnum      ' Screen the user is currently on:  encounter, template edit, etc
Private mlngTemplateId As Long           ' Current Template Id (encounter or SO), if any

Public Property Get SnoID() As Long
    SnoID = mlngSnoID
End Property
Public Property Let SnoID(ByVal NewValue As Long)
    mlngSnoID = NewValue
End Property

Public Property Get ViewMode() As Integer
    ViewMode = mintViewMode
End Property
Public Property Let ViewMode(ByVal NewValue As Integer)
    mintViewMode = NewValue
End Property

Public Property Get EntryTab() As Integer
    EntryTab = mintEntryTab
End Property
Public Property Let EntryTab(ByVal NewValue As Integer)
    mintEntryTab = NewValue
End Property

Public Property Get ListSize() As Integer
    ListSize = mintListsize
End Property
Public Property Let ListSize(ByVal NewValue As Integer)
    mintListsize = NewValue
End Property

Public Property Get Sarg() As String
    Sarg = mstrSarg
End Property
Public Property Let Sarg(ByVal NewValue As String)
    mstrSarg = NewValue
End Property

Public Property Get TemplateItems() As Collection
    Set TemplateItems = mcolTemplateItems
End Property
Public Property Set TemplateItems(ByRef NewValue As Collection)
    Set mcolTemplateItems = NewValue
End Property

Public Property Get ScreenMode() As ModeEnum
    ScreenMode = menumScreenMode
End Property
Public Property Let ScreenMode(ByVal NewValue As ModeEnum)
    menumScreenMode = NewValue
End Property

Public Property Get TemplateId() As Long
    TemplateId = mlngTemplateId
End Property
Public Property Let TemplateId(ByVal NewValue As Long)
    mlngTemplateId = NewValue
End Property
