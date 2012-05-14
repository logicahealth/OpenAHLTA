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
Attribute VB_Name = "AllergySearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=AllergySearchParam
'--- 1/22/2004 10:49:54 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iParameters
Implements iObjectPersist

Private Const OBJECT_VERSION As String = "1.0"
Private Const UNIQUE_IDENTIFIER As String = "AllergySearchParam"

Private mobjParams As iParameters

Private mIncludeReactions As Boolean                'should reactions be included in the results?
Private mIncludeComment As Boolean                  'should comments be included in the results?
Private mPatientUnitNumber As String                'patient unit number
Private mPatientIENNumber As String                 'patient ien number
Private mTimeObject_TimeSearchType As Integer       'TimeObject properties
Private mTimeObject_LastN As Integer                '       "
Private mTimeObject_CalcSlidingTime As Double       '       "
Private mTimeObject_BeginTime As Double             '       "
Private mTimeObject_EndTime As Double               '       "
Private mTimeObject_EncounterNumber As String       '       "
Private mTimeObject_EncounterFacilityNCID As String '       "


Private Const cINCLUDE_REACTIONS As String = "IncludeReactions"
Private Const cINCLUDE_COMMENT As String = "IncludeComment"
Private Const cPATIENT_UNIT_NUMBER As String = "PatientUnitNumber"
Private Const cPATIENT_IEN_NUMBER As String = "PatientIENNumber"
Private Const cTIMEOBJECT_TIMESEARCHTYPE As String = "TimeSearchType"
Private Const cTIMEOBJECT_LASTN As String = "TimeLastN"
Private Const cTIMEOBJECT_CALCSLIDINGTIME As String = "TimeCalcSlidingTime"
Private Const cTIMEOBJECT_BEGINTIME As String = "TimeBeginTime"
Private Const cTIMEOBJECT_ENDTIME As String = "TimeEndTime"
Private Const cTIMEOBJECT_ENCOUNTERNUMBER As String = "TimeEncounterNumber"
Private Const cTIMEOBJECT_ENCOUNTERFACILITYNCID As String = "TimeEncountFacilityNCID"


'--- end init / terminate procedures


'--- begin property lets/gets
Public Property Get IncludeReactions() As Boolean
    IncludeReactions = GetBoolValue(cINCLUDE_REACTIONS)
End Property

Public Property Let IncludeReactions(ByVal Value As Boolean)
    SetBoolValue cINCLUDE_REACTIONS, Value
End Property

Public Property Get IncludeComment() As Boolean
    IncludeComment = GetBoolValue(cINCLUDE_COMMENT)
End Property

Public Property Let IncludeComment(ByVal Value As Boolean)
    SetBoolValue cINCLUDE_COMMENT, Value
End Property

Public Property Get PatientUnitNumber() As String
    PatientUnitNumber = GetStringValue(cPATIENT_UNIT_NUMBER)
End Property

Public Property Let PatientUnitNumber(ByVal Value As String)
    SetStringValue cPATIENT_UNIT_NUMBER, Value
End Property

Public Property Get PatientIENNumber() As String
    PatientIENNumber = GetStringValue(cPATIENT_IEN_NUMBER)
End Property

Public Property Let PatientIENNumber(ByVal Value As String)
    SetStringValue cPATIENT_IEN_NUMBER, Value
End Property

Public Property Get TimeObject_TimeSearchType() As Integer
    TimeObject_TimeSearchType = GetIntValue(cTIMEOBJECT_TIMESEARCHTYPE)
End Property

Public Property Let TimeObject_TimeSearchType(ByVal Value As Integer)
    SetIntValue cTIMEOBJECT_TIMESEARCHTYPE, Value
End Property

Public Property Get TimeObject_LastN() As Integer
    TimeObject_LastN = GetIntValue(cTIMEOBJECT_LASTN)
End Property

Public Property Let TimeObject_LastN(ByVal Value As Integer)
    SetIntValue cTIMEOBJECT_LASTN, Value
End Property

Public Property Get TimeObject_CalcSlidingTime() As Double
    TimeObject_CalcSlidingTime = GetDoubleValue(cTIMEOBJECT_CALCSLIDINGTIME)
End Property

Public Property Let TimeObject_CalcSlidingTime(ByVal Value As Double)
    SetDoubleValue cTIMEOBJECT_CALCSLIDINGTIME, Value
End Property

Public Property Get TimeObject_BeginTime() As Double
    TimeObject_BeginTime = GetDoubleValue(cTIMEOBJECT_BEGINTIME)
End Property

Public Property Let TimeObject_BeginTime(ByVal Value As Double)
    SetDoubleValue cTIMEOBJECT_BEGINTIME, Value
End Property

Public Property Get TimeObject_EndTime() As Double
    TimeObject_EndTime = GetDoubleValue(cTIMEOBJECT_ENDTIME)
End Property

Public Property Let TimeObject_EndTime(ByVal Value As Double)
    SetDoubleValue cTIMEOBJECT_ENDTIME, Value
End Property
Public Property Get TimeObject_EncounterNumber() As String
    TimeObject_EncounterNumber = GetStringValue(cTIMEOBJECT_ENCOUNTERNUMBER)
End Property

Public Property Let TimeObject_EncounterNumber(ByVal Value As String)
    SetStringValue cTIMEOBJECT_ENCOUNTERNUMBER, Value
End Property

Public Property Get TimeObject_EncounterFacilityNCID() As String
    TimeObject_EncounterFacilityNCID = GetStringValue(cTIMEOBJECT_ENCOUNTERFACILITYNCID)
End Property

Public Property Let TimeObject_EncounterFacilityNCID(ByVal Value As String)
    SetStringValue cTIMEOBJECT_ENCOUNTERFACILITYNCID, Value
End Property


'--- end property lets/gets

Private Sub Class_Initialize()
    Set mobjParams = New ParametersGeneric
    mTimeObject_TimeSearchType = -999
End Sub

Private Function iObjectPersist_Copy() As Object
    Dim persist As iObjectPersist

    Set persist = mobjParams
    Set iObjectPersist_Copy = persist.Copy
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
    Dim persist As iObjectPersist
    
    Set persist = mobjParams
    persist.persist Writer
End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
    Dim persist As iObjectPersist

    Set persist = mobjParams
    iObjectPersist_Restore = persist.Restore(Reader)
End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
    Dim persist As iObjectPersist

    Set persist = mobjParams
    iObjectPersist_UniqueIdentifier = persist.UniqueIdentifier
End Property

Private Property Get iObjectPersist_Version() As String
    Dim persist As iObjectPersist

    Set persist = mobjParams
    iObjectPersist_Version = persist.Version
End Property

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As DL_Support.Parameter
    mobjParams.Add Name, Value
End Function

Private Function iParameters_Count() As Long
    iParameters_Count = mobjParams.count
End Function

Private Function iParameters_Exists(ByVal Name As String) As Boolean
    iParameters_Exists = mobjParams.Exists(Name)
End Function

Private Function iParameters_Item(ByVal Key As Variant) As DL_Support.Parameter
    Set iParameters_Item = mobjParams.Item(Key)
End Function

Private Sub iParameters_Remove(ByVal Key As Variant)
    mobjParams.Remove (Key)
End Sub

Private Function GetBoolValue(ByVal Name As String) As Boolean
  If mobjParams.Exists(Name) Then
    GetBoolValue = mobjParams.Item(Name).Value
  Else
    GetBoolValue = False
  End If
End Function

Private Sub SetBoolValue(ByVal Name As String, ByVal NewValue As Boolean)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub

Private Function GetIntValue(ByVal Name As String) As Integer
  If mobjParams.Exists(Name) Then
    GetIntValue = mobjParams.Item(Name).Value
  Else
    GetIntValue = 0
  End If
End Function

Private Sub SetIntValue(ByVal Name As String, ByVal NewValue As Integer)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub

Private Function GetStringValue(ByVal Name As String) As String
  If mobjParams.Exists(Name) Then
    GetStringValue = mobjParams.Item(Name).Value
  Else
    GetStringValue = vbNullString
  End If
End Function

Private Sub SetStringValue(ByVal Name As String, ByVal NewValue As String)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub

Private Function GetDoubleValue(ByVal Name As String) As Double
  If mobjParams.Exists(Name) Then
    GetDoubleValue = CDbl(mobjParams.Item(Name).Value)
  Else
    GetDoubleValue = 0
  End If
End Function

Private Sub SetDoubleValue(ByVal Name As String, ByVal NewValue As Double)
  If mobjParams.Exists(Name) Then
    mobjParams.Item(Name).Value = NewValue
  Else
    mobjParams.Add Name, NewValue
  End If
End Sub

