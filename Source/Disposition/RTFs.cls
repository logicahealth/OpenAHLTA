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
Attribute VB_Name = "RTFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Disposition As String
Public InjuryIllness As String
Public FollowUp As String
Public Discussed As String
Public TimeFactor As String

Public AdminCodes As String
Public BillingAdmin As String
Public EM As String
Public WorkStatus As String

Public XtraEM1 As String
Public XtraEM2 As String
Public XtraEM3 As String

Public DentalStatus As String

'SCR 65426
Public Discharge As String

Public Function RTFNote() As String
    Dim sz As String
    Dim objEnc As CHCSIIEncounterOps.Encounter
    
    sz = Disposition
    
    If Len(FollowUp) Then
        sz = sz & RTF_NEWLINE & FollowUp
    End If
    
    If Len(Discussed) Then
        sz = sz & RTF_NEWLINE & Discussed
    End If
    
    'IN PGUI Mode do not display if HIPAA switch is on SCR-50455
    If gobjShared.IsAppMode(modeTheater) Or (Len(gobjShared.CmdLineSwitches("HIPAA837")) = 0 And gobjShared.IsAppMode(modeCHCSI_GUI)) Then   '<SCR 36876
        If Len(InjuryIllness) Then
            sz = sz & RTF_NEWLINE & InjuryIllness
        End If
    End If
    
    
    
    If Len(AdminCodes) Then
        sz = sz & RTF_NEWLINE & AdminCodes
    End If
    
    If Len(BillingAdmin) Then
        sz = sz & RTF_NEWLINE & BillingAdmin
    End If
   
'    Do not include in Disposition RTF SCR-35183
    'Include E&M Code if Option Selected SCR-38961
    Set objEnc = gobjEncParent.CurrentEncounter
    If objEnc.EncProperties.IncludeCodes Or gobjShared.IsAppMode(modeTheater) Then
        If Len(EM) Then
            sz = sz & RTF_NEWLINE & EM
        End If
    
        If Len(XtraEM1) Then
            sz = sz & RTF_NEWLINE & XtraEM1
        End If
        If Len(XtraEM2) Then
            sz = sz & RTF_NEWLINE & XtraEM2
        End If
        If Len(XtraEM3) Then
            sz = sz & RTF_NEWLINE & XtraEM3
        End If
    End If
    
   If Len(TimeFactor) Then
        sz = sz & RTF_NEWLINE & TimeFactor
    End If
 
'   Do not place workstatus in Disposition RTF.  Workstatus changes are displayed in Add Note Section
'   b/c workstatus can be changed in Screening and Patient Demographics
'    If Len(WorkStatus) Then
'        sz = sz & RTF_NEWLINE & WorkStatus
'    End If
    
    'KDunne 1/29/2003
    If Len(DentalStatus) > 0 Then
        sz = sz & RTF_NEWLINE & DentalStatus
    End If
    
    'SF 2-08-05 SCR 65426
    If Len(Discharge) > 0 Then
        sz = sz & RTF_NEWLINE & Discharge
    End If
    
    'Remove Leading Space
    If InStr(sz, RTF_NEWLINE) = 1 Then
        sz = Right(sz, Len(sz) - Len(RTF_NEWLINE))
    End If
    
    'Added line space SCR-33921
    RTFNote = "\fs16 " & sz & RTF_NEWLINE
    
End Function
