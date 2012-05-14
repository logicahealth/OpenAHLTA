VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimingType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarUnexpandedTime As String 'local copy
Private mvarAdminTime As String 'local copy
Private mvarFrequency As String 'local copy
Private mvarDuration As Single  'local copy
Private mvarStartDate As String  'local copy
Private mvarStopDate As String  'local copy
'Private mvarOnexOrContinued As String 'local copy
Private mvarOnexOrContinued As hlRADLAB 'local copy
Private mvarStartTime As String 'local copy
Private mvarStopTime As String  'local copy
Private hlDurationUnits As hlDuration

Public Enum hlDuration
  hlDuration_Seconds = 0
  hlDuration_Minutes
  hlDuration_Hours
  hlDuration_Days
  hlDuration_Weeks
  hlDuration_Months
  hlDuration_Xinterval
  hlDuration_Total
End Enum
  
Friend Property Let StopTime(ByVal vData As String)
    mvarStopTime = vData
End Property

Public Property Get StopTime() As String
    StopTime = mvarStopTime
End Property

Public Property Let StartTime(ByVal vData As String)
'Friend Property Let StartTime(ByVal vData As String)

    mvarStartTime = vData
End Property

Public Property Get StartTime() As String

    StartTime = mvarStartTime
    
End Property

Public Property Let OnexOrContinued(ByVal vData As hlRADLAB)
    mvarOnexOrContinued = vData
End Property

Public Property Get OnexOrContinued() As hlRADLAB
    OnexOrContinued = mvarOnexOrContinued
End Property


Public Property Let StopDate(ByVal szData As String)
     On Error GoTo ErrorHandler:
     If IsDate(szData) Then
          mvarStopDate = Format$(szData, "yyyymmdd")
     Else
          mvarStopDate = szData
     End If
     Exit Property
ErrorHandler:
     mvarStopDate = szData
End Property

Public Property Get StopDate() As String
    StopDate = mvarStopDate
End Property

Public Property Let StartDate(ByVal szData As String)
     On Error GoTo ErrorHandler:
     If IsDate(szData) Then
          mvarStartDate = Format$(szData, "yyyymmdd")
     Else
          mvarStartDate = szData
     End If
     Exit Property
ErrorHandler:
     mvarStartDate = szData
End Property

Public Property Get StartDate() As String
    StartDate = mvarStartDate
End Property

Public Property Let Duration(ByVal vData As Single)
    mvarDuration = vData
End Property

Public Property Get Duration() As Single
    Duration = mvarDuration
End Property

Public Property Let Frequency(ByVal vData As String)
    mvarFrequency = vData
End Property

Public Property Get Frequency() As String
    Frequency = mvarFrequency
End Property

Public Property Let AdminTime(ByVal vData As String)
    mvarAdminTime = vData
End Property

Public Property Get AdminTime() As String
    AdminTime = mvarAdminTime
End Property

Public Property Let UnexpandedTime(ByVal vData As String)
    mvarUnexpandedTime = vData
End Property

Public Property Get UnexpandedTime() As String
    UnexpandedTime = mvarUnexpandedTime
End Property
'Friend Property Let DurationUnits(hlData As hlDuration)
''Currently CHCS only uses duration units of days. The hl7 chcs icd lists several other
'  'units, but they are not currently implemented.
'  'Bug in CHCS- accepts the letters other than D, without any warning or error
'  'but places the order as units of days.
'  'This property will be initialized to Days, and will not allow changes unless CHCS decides to
'  'implement the other units.
'  hlDurationUnits = hlData
'End Property
Public Property Get DurationUnits() As hlDuration
  
  DurationUnits = hlDurationUnits
End Property

Private Sub Class_Initialize()
  'Initialize some properties
  hlDurationUnits = hlDuration_Days
  mvarDuration = 0
  mvarOnexOrContinued = hlONETIME
End Sub
