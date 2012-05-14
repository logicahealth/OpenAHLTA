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
Attribute VB_Name = "NonProviderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrFacilityNCID As String
Private mstrNursingDivNCID As String
Private mstrUserNCID As String
Private mstrEncProviderNCID As String
Private mobjConfig As Config
Private mblnUserIsSignable As Boolean

Public Property Get IsSignable() As Boolean
    IsSignable = mblnUserIsSignable
End Property

Public Function SearchClinician() As dtos.Providers
Dim objClinician As Clinician
Dim objProvs As dtos.Providers

    Set objClinician = New Clinician
    
    'Set Clinician to look for Only Providers (ClinicClient.vbp; ClinOps.cls modification)
    objClinician.Type_Of_Search = SearchType.Providers
    
    'Since we modified ClinOps in the ClinicClient.vbp, we can search in the normal fashion
    If objClinician.SearchClinician(mstrFacilityNCID, "", cwobjLogonEnviron, True, False) Then
    
        Set objProvs = RetrieveProviders(objClinician.NCID)
    End If

    Set SearchClinician = objProvs
End Function

Public Property Set ConfigObject(objConfig As Config)
    Set mobjConfig = objConfig
End Property

Public Function getProviders() As dtos.Providers
    Dim objProviders As dtos.Providers
    
    Set objProviders = RetrieveProviders()
    
    Set getProviders = objProviders
End Function

Public Function OpenNonProvider(robjLogon As Object, robjComm As Object, robjUser As Object) As dtos.Providers
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim objCurrentEncounter As CHCSIIEncounterOps.Encounter
    Dim objProviders As dtos.Providers
    
    Dim strProviderNCID As String
    
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogon

    Set objEncParent = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    If Not objEncParent Is Nothing Then
        Set objCurrentEncounter = objEncParent.CurrentEncounter
        mstrEncProviderNCID = objCurrentEncounter.PrimaryProvider.NCID
    End If
    'We're done with them; lose 'em.
    Set objCurrentEncounter = Nothing
    Set objEncParent = Nothing
    
    mstrNursingDivNCID = robjLogon.NursingDivisionNCID
    mstrUserNCID = cwobjLogonEnviron.UserNCID
    mstrFacilityNCID = robjLogon.FacilityNCID
    
    Set objProviders = RetrieveProviders(mstrUserNCID)
    
    mblnUserIsSignable = False
    'If user can't sign, add the Encounter Provider to the list
    If objProviders.Exists(mstrUserNCID) Then
        If Not objProviders(mstrUserNCID).IsSignable Then
            Set objProviders = RetrieveProviders(mstrEncProviderNCID)
        Else
            mblnUserIsSignable = True
        End If
    End If
    
    Set OpenNonProvider = objProviders
End Function

Private Sub Class_Initialize()
'  Debug.Print "CHCSII_NonProvider Class_Initialize"
  
#If coiDebug Then
     gCOI_DL.debugStart
     DR.TimerTime = 0
     gbDebugStarted = True
#End If

    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared

    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
End Sub

Private Function RetrieveProviders(Optional ClinicianNCID As String = "", Optional ProviderIEN As String = "") As dtos.Providers

Dim objSearchParams As DM_NonProviderList.SearchNPParams
Dim objDMProv As DL_Interfaces.iMapperNonProviderList
Dim objOshMan As OshMan

    Set objSearchParams = New DM_NonProviderList.SearchNPParams
    
    objSearchParams.NursingDivisionNCID = mstrNursingDivNCID
    objSearchParams.FacilityNCID = mstrFacilityNCID
    objSearchParams.UserNCID = mstrUserNCID
    objSearchParams.ClinicianNCID = ClinicianNCID
    objSearchParams.ProviderIEN = ProviderIEN
    
    Set objDMProv = gobjShared.DataAccessService.NonProviderList

    'Set objOshMan = New OshMan
    'objOshMan.ConfigObject = mobjConfig
    'Set objDMProv.NPOEOshMan = objOshMan
    Set RetrieveProviders = objDMProv.Selected(objSearchParams)

End Function

Public Sub SaveList()
Dim objSearchParams As DM_NonProviderList.SearchNPParams
Dim objDMProv As DL_Interfaces.iMapperNonProviderList

    Set objSearchParams = New DM_NonProviderList.SearchNPParams
    
    'SaveList only needs the UserNCID, and ListMode parameter
    With objSearchParams
        .UserNCID = mstrUserNCID
        .ListMode = Provider
    End With

    Set objDMProv = gobjShared.DataAccessService.NonProviderList
    Call objDMProv.SaveList(objSearchParams)

End Sub
