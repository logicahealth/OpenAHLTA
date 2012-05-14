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
Attribute VB_Name = "FormDD741"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormDD741.cls"
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics

Public Function BuildDD741(sPatientUnitNumber As String, _
                           oReport As CRAXDRT.Report) As Boolean
On Error GoTo ErrHandler

BuildDD741 = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As String
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    '-- Patient Demographics
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "demographics"

    '-- Sponsor Demographics
    '----------------------------------------------------------
    If oRec.Fields.Item("FMP").Value = "20" Then
        sSponsorUnitNumber = sPatientUnitNumber
    Else
        sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
    End If
    If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "sponsor"
    
    '-- Patient Address
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetAddress(sPatientUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "address"

'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"
    
'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("DD741.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp

BuildDD741 = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".BuildDD741", App.EXEName, vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
End Function

Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub

Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
End Sub




