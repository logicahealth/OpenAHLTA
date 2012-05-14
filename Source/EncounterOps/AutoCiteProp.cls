VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoCiteProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Set of user preferences for autociting on the encounter doc"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mbSelected As Integer         '- is this guy selected?
Private mvarAppNCID As Long           '- The app NCID for this AC module
Private mvarAppName As String         '- the name of the ole server
Private mvarNumberPeriods As Integer  '- number of periods
Private mvarPeriod As AC_Period       '- End of the Inclusion range
Private mvarNumberOfEvents As Long    '- Number of past events
Private mvarSensitivityLevel As Long  '- Sensitivity level to return.
Private sModule As String

' 7/08/2002
'WR moved to Encops_itf.tlb SCR 25225
'Public Enum AC_Period
'  Hour = 1
'  Day = 2
'  Month = 3
'  Year = 4
'End Enum

Public Property Get AppName() As String:                   AppName = mvarAppName:                   End Property
Public Property Let AppName(RHV As String):                mvarAppName = RHV:                       End Property
Public Property Get AppNCID() As Long:                     AppNCID = mvarAppNCID:                   End Property
Public Property Let AppNCID(ByVal vData As Long):          mvarAppNCID = vData:                     End Property
Public Property Get NumberOfEvents() As Long:              NumberOfEvents = mvarNumberOfEvents:     End Property
Public Property Let NumberOfEvents(ByVal vData As Long):   mvarNumberOfEvents = vData:              End Property
Public Property Get NumberPeriods() As Long:               NumberPeriods = mvarNumberPeriods:       End Property
Public Property Let NumberPeriods(ByVal vData As Long):    mvarNumberPeriods = vData:               End Property
Public Property Get Period() As Long:                      Period = mvarPeriod:                     End Property
Public Property Let Period(ByVal vData As Long):           mvarPeriod = vData:                      End Property
Public Property Get Selected() As Integer:                 Selected = mbSelected:                   End Property
Public Property Let Selected(RHV As Integer):              mbSelected = RHV:                        End Property
Public Property Get SensitivityLevel() As Long:            SensitivityLevel = mvarSensitivityLevel: End Property
Public Property Let SensitivityLevel(ByVal vData As Long): mvarSensitivityLevel = vData:            End Property


