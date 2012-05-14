VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateHxSOItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Data Structure for storing user entry history for GoBack and go forward
Private mlSnoID As String             ' Current SnoID at change
Private mlTemplateId As Long           ' Current Template Id (encounter or SO), if any
Private miViewMode As Integer          ' Current view mode
Private miEntryTab As Integer          ' Current entry tab
Private miListsize As Integer          ' Current list size
Private msSarg As String               ' search string, key array, or Dx array
Private mcolCCol As Collection         ' Template collection, if applicable
Private meScreenMode As ModeEnum       ' Screen the user is currently on:  encounter, template edit, etc


Public Enum ModeEnum
    Encmode = 0
    TmplSelectMode
    TmplEditMode
    TmplAddinMode
    FormMode
End Enum

Public Property Get ViewMode() As Integer
   ViewMode = miViewMode
End Property
Public Property Let ViewMode(val As Integer)
    miViewMode = val
End Property
Public Property Get EntryTab() As Integer
    EntryTab = miEntryTab
End Property
Public Property Let EntryTab(val As Integer)
    miEntryTab = val
End Property
Public Property Get Listsize() As Integer
    Listsize = miListsize
End Property
Public Property Let Listsize(val As Integer)
    miListsize = val
End Property
Public Property Get SnoID() As Long
    SnoID = mlSnoID
End Property
Public Property Let SnoID(val As Long)
    mlSnoID = val
End Property
Public Property Get TemplateID() As Long
    TemplateID = mlTemplateId
End Property
Public Property Let TemplateID(val As Long)
    mlTemplateId = val
End Property
Public Property Get Sarg() As String
    Sarg = msSarg
End Property
Public Property Let Sarg(val As String)
    msSarg = val
End Property
Public Property Get CCol() As Collection
    Set CCol = mcolCCol
End Property
Public Property Set CCol(val As Collection)
    If Not val Is Nothing Then Set mcolCCol = val
End Property
Public Property Get ScreenMode() As ModeEnum
    ScreenMode = meScreenMode
End Property
Public Property Let ScreenMode(val As ModeEnum)
    meScreenMode = val
End Property

