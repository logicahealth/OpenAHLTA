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
Attribute VB_Name = "Encounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements IPersist

Private Const HEADER As String = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
            & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
            & "{\colortbl;\red0\green0\blue0;}" _
            & "\pard\plain\f0\fs16\b\ul %1 \b0 %2 by %3 @ %4\par " _
            & "\pard\plain \widctlpar\adjustright \cgrid %5}"

Private Const NEW_NOTE As String = "\par "

Public RELATED_TO_INPATIENT As String

Private mvarPATIENT_STATUS As String 'local copy
Private mvarWORK_LOAD_INDICATOR As String 'local copy
Private mvarENCOUNTER_STATUS As String 'local copy

Private mvarRTFs As RTFs

Private mvarAnPStatus As Long
Private mvarDispoStatus As Long

Private mvarValidator As Validation

Public Property Get AnPStatus() As Long
  AnPStatus = mvarAnPStatus
End Property

Public Property Let AnPStatus(ByVal NewValue As Long)
  mvarAnPStatus = NewValue
End Property

Public Property Get DispositionStatus() As Long
  DispositionStatus = mvarDispoStatus
End Property

Public Property Let DispositionStatus(ByVal NewValue As Long)
  mvarDispoStatus = NewValue
End Property

Public Property Get RTFs() As RTFs
  Set RTFs = mvarRTFs
End Property

Public Property Let RTFs(ByVal NewValue As RTFs)
  Set mvarRTFs = NewValue
End Property

Public Property Let WORK_LOAD_INDICATOR(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WORK_LOAD_INDICATOR = 5
    
    mvarValidator.ValidatePropByList "Encounter", "WORK_LOAD_INDICATOR", vData, "C", "N", vbNullString
    mvarWORK_LOAD_INDICATOR = vData
    
End Property


Public Property Get WORK_LOAD_INDICATOR() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WORK_LOAD_INDICATOR
    WORK_LOAD_INDICATOR = UCase(mvarWORK_LOAD_INDICATOR)
End Property

Public Property Let Patient_Status(ByVal vData As String)

    mvarValidator.ValidatePropByList "Encounter", "PATIENT_STATUS", vData, "OUTPATIENT", "INPATIENT", "Outpatient Observation", "Outpatient APV"
    mvarPATIENT_STATUS = vData
    
End Property

Public Property Get Patient_Status() As String
    Patient_Status = mvarPATIENT_STATUS
End Property

Public Property Let ENCOUNTER_STATUS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ENCOUNTER_STATUS = 5
    mvarENCOUNTER_STATUS = vData
    'msgbox "Validate Encounter Status"
End Property

Public Property Get ENCOUNTER_STATUS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ENCOUNTER_STATUS
    ENCOUNTER_STATUS = mvarENCOUNTER_STATUS
End Property

Public Sub AppendNote(ByVal NewNote As String)
  With mvarRTFs.Item(4)
    If Len(.RTF) > 0 Then
      .RTF = .RTF & NEW_NOTE & NewNote
    Else
      .RTF = NewNote
    End If
  End With
End Sub

Public Function CreateNote(ByVal CreatedBy As String) As String
Dim SoRTF As String
Dim AnPRTF As String
Dim DispositionRTF As String
Dim NoteRTF As String

  If Len(mvarRTFs.Item(3).RTF) > 0 Then
    SoRTF = Replace(SoRTF, "%5", mvarRTFs.Item(3).RTF)
  End If
  
  If Len(mvarRTFs.Item(1).RTF) > 0 Then
    AnPRTF = Replace(HEADER, "%1", mvarRTFs.Item(1).Title)
    AnPRTF = Replace(AnPRTF, "%2", mvarRTFs.Item(1).Verb)
    AnPRTF = Replace(AnPRTF, "%3", mvarRTFs.Item(1).WrittenBy)
    AnPRTF = Replace(AnPRTF, "%4", mvarRTFs.Item(1).OnsetDate)
    AnPRTF = Replace(AnPRTF, "%5", mvarRTFs.Item(1).RTF)
  End If
  
  If Len(mvarRTFs.Item(2).RTF) > 0 Then
    DispositionRTF = Replace(HEADER, "%1", mvarRTFs.Item(2).Title)
    DispositionRTF = Replace(DispositionRTF, "%2", mvarRTFs.Item(2).Verb)
    DispositionRTF = Replace(DispositionRTF, "%3", mvarRTFs.Item(2).WrittenBy)
    DispositionRTF = Replace(DispositionRTF, "%4", mvarRTFs.Item(2).OnsetDate)
    DispositionRTF = Replace(DispositionRTF, "%5", mvarRTFs.Item(2).RTF)
  End If
  
  NoteRTF = Replace(HEADER, "%1", "Note")
  NoteRTF = Replace(NoteRTF, "%2", "CreatedBy")
  NoteRTF = Replace(NoteRTF, "%3", CreatedBy)
  NoteRTF = Replace(NoteRTF, "%4", Format$(Now, "dd mmm yyyy hhnn"))
  NoteRTF = Replace(NoteRTF, "%5", SoRTF & NEW_NOTE & AnPRTF & NEW_NOTE & DispositionRTF)
  
  CreateNote = NoteRTF

End Function

Public Function CurrentNote() As String
  On Error Resume Next
  CurrentNote = mvarRTFs.Item(4).RTF
End Function

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
    WORK_LOAD_INDICATOR = "NOCOUNT"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mvarRTFs = New RTFs
  
    mvarRTFs.Add "A/P"
    mvarRTFs.Add "Disposition"
    mvarRTFs.Add "SO"
    mvarRTFs.Add "Note"
End Sub

Private Sub Class_Terminate()
    Set mvarValidator = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Persistence Block
''''''''''''''''''''''''''''''''''''''''''''''
Private Function IPersist_ChildFactory() As Object

End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
End Function

