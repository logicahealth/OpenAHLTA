VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This holds property fields used by the demo
Public Type propertyRecord
  SnoID As String
  nodeKey As String
  Flags As Integer
  eFlags As Integer
  qFlags As Integer           ' Qualifier flag
  CAflags As Integer           ' Children Attribute flags
  termtype As Byte
  rxCode As String '* 1
  Cpt As String               ' cpt from codelinks (here for convenience)
  Icd9cm As String            ' icd from codelinks ( ditto )
  Unit As String
  fmtUnit As String           ' Used to hold a format for value
  RangeNormalHigh As Double
  RangeNormalLow As Double
End Type


'
' This data structure is used to store the data associated
' with each finding entered into an encounter
'
Public Type encounterrecord   ' For each finding in an encounter
  ItemId As Long              ' Assigned by Medcin control
  UniqueId As Long            ' Used for the itemtag
  eTime As Date               ' Encounter date
  Chart As ListToolRecord     ' The chart record structure
  Prec As propertyRecord      ' The property record structure
  Text As String              ' Copy of displayed text
  Inactive As Boolean         ' set True if record removed from encounter
  Group As Integer            ' Group number calc'd by Medcin
  KeyLen As Integer           ' length of nodekey
  Positive As Boolean         ' True if pos button, else false
  MergeRequired As Boolean    ''SCR 23945 Indicates problems merging multiple notes, requires user action
  MergeNoteNew As String      ''SCR 23945 Note text that needs to be merged by the user
  MergeNoteCurr As String     ''SCR 23945 The note that has already been entered
End Type

