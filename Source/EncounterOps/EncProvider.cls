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
Attribute VB_Name = "EncProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An Encounter contains a collection of these"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Const MODULE_NAME = "EncProvider"

Public Enum enumRecordState
  rsUnknown
  rsNew
  rsDelete
  rsRemove
  rsSaved
  rsModified
End Enum

Private msIEN As String
Private msNCID As Long
Private msFullName As String
Private msFacility As Variant
Private msFacilityNCID As Variant
Private msRoleNCID As String
Private msRole As String
Private msUnitNumber As String
Private msSigLine2 As String
Private msSigLine3 As String
Private mcProcedures As Collection

Private meCurrentState As enumRecordState
Private msPrevNCID As String

Public Property Get Facility() As Variant:             Facility = msFacility:          End Property
Public Property Let Facility(ByVal vData As Variant):  msFacility = vData:             End Property
Public Property Set Facility(ByVal vData As Variant):  msFacility = vData:             End Property
Public Property Get FacilityNCID() As String:          FacilityNCID = msFacilityNCID:  End Property
Public Property Let FacilityNCID(RHV As String):       msFacilityNCID = RHV:           End Property
Public Property Get FullName() As String:              FullName = msFullName:          End Property
Public Property Let FullName(ByVal vData As String):   msFullName = vData:             End Property
Public Property Get IEN() As String:                   IEN = msIEN:                    End Property
Public Property Let IEN(ByVal sData As String):        msIEN = sData:                  End Property
Public Property Get NCID() As Long:                    NCID = msNCID:                  End Property
Public Property Let NCID(ByVal vData As Long):         msNCID = vData:                 End Property
Public Property Get Role() As String:                  Role = msRole:                  End Property
Public Property Let Role(RHV As String):               msRole = RHV:                   End Property
Public Property Get RoleNCID() As String:              RoleNCID = msRoleNCID:          End Property
Public Property Let RoleNCID(RHV As String):           msRoleNCID = RHV:               End Property
Public Property Get UnitNumber() As String:            UnitNumber = msUnitNumber:      End Property
Public Property Let UnitNumber(RHV As String):         msUnitNumber = RHV:             End Property

Friend Property Get SigLine2() As String
    SigLine2 = msSigLine2
End Property

Friend Property Let SigLine2(ByRef RHV As String)
    msSigLine2 = RHV
End Property

Friend Property Get SigLine3() As String
    SigLine3 = msSigLine3
End Property

Friend Property Let SigLine3(ByRef RHV As String)
    msSigLine3 = RHV
End Property

Private Function FindProcedure(ByVal sDataId As String, ByVal sSnoID As String) As EncProviderProcedure

Dim oProc As EncProviderProcedure
Dim lIdx As Long

   If mcProcedures Is Nothing Then
      Set FindProcedure = Nothing
      Exit Function
   End If
   
   'Make sure this is not a duplicate - Compare DataId if supplied or sSnoID
   For lIdx = 1 To mcProcedures.Count
      If Len(sDataId) > 0 And sDataId = mcProcedures.Item(lIdx).DataId Then
         Set FindProcedure = mcProcedures.Item(lIdx)
         Exit For
      ElseIf Len(sSnoID) > 0 And sSnoID = mcProcedures.Item(lIdx).SnoID Then
         Set FindProcedure = mcProcedures.Item(lIdx)
         Exit For
      End If
   Next
   
End Function

Public Sub UnAssociateProcedure(ByVal sDataId As String, ByVal sSnoID As String)

On Error GoTo ErrorHandler

Dim oProc As EncProviderProcedure

   'Get existing procedure
   Set oProc = FindProcedure(sDataId, sSnoID)
   
   If Not oProc Is Nothing Then
      If oProc.State = rsNew Then
         'Remove it from the collection
         oProc.State = rsRemove
         RemoveProcedures
      Else
         oProc.State = rsDelete
      End If
   End If
   
ErrorHandler:
   Set oProc = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "UnAssociateProcedure"
   
End Sub

Public Sub UpdateProcedure(ByVal sDataId As String, ByVal sSnoID As String)

On Error GoTo ErrorHandler

Dim oProc As EncProviderProcedure
   
   'Get existing procedure
   Set oProc = FindProcedure(sDataId, sSnoID)
   If Not oProc Is Nothing Then oProc.DataId = sDataId
     
ErrorHandler:
   Set oProc = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "UnAssociateProcedure"
        
End Sub
Public Sub AssociateProcedure(ByVal sDataId As String, ByVal sSnoID As String)

On Error GoTo ErrorHandler

Dim oProc As EncProviderProcedure
Dim lIdx As Long
   
   If mcProcedures Is Nothing Then
      Set mcProcedures = New Collection
   End If
   
   'Get existing procedure
   Set oProc = FindProcedure(sDataId, sSnoID)
      
   If oProc Is Nothing Then
      'Add new procedure
      Set oProc = New EncProviderProcedure
      oProc.Init sDataId, sSnoID
      mcProcedures.Add oProc
   Else
      'Update existing procedure
      If oProc.DataId <> sDataId Or oProc.State = rsDelete Then
         oProc.DataId = sDataId
         oProc.State = rsModified
      End If
      oProc.SnoID = sSnoID
   End If
   
ErrorHandler:
   Set oProc = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "AssociateProcedure"
   
End Sub

Public Function HasProcedure(ByVal sDataId As String, sSnoID As String) As Boolean
   
Dim oProvider As EncProviderProcedure

   Set oProvider = FindProcedure(sDataId, sSnoID)
   If Not oProvider Is Nothing Then
      HasProcedure = (oProvider.State = rsNew Or oProvider.State = rsSaved Or oProvider.State = rsModified)
   Else
      HasProcedure = False
   End If
   
End Function

Private Sub Class_Terminate()
   Set mcProcedures = Nothing
End Sub

Friend Function CDRLoad(ByRef oClinicalObservation As ClinicalObservation) As Boolean

On Error GoTo ErrorHandler

Dim lModIdx          As Long
Dim oObsMod          As ClinicalObservation
Dim oProcedure       As EncProviderProcedure

   With oClinicalObservation
         
      NCID = .ObsValue.value.CodedWOSform.ValueConcept
      FullName = .ObsValue.value.CodedWOSform.ValueConcept.PrefRep("2000").Representation

      For lModIdx = 1 To .ObsMods.Count
         
         'Don't use for each because of memory leak with 3M objects
         Set oObsMod = .ObsMods(lModIdx)
         
         'Process each mod value
         Select Case oObsMod.ObsId.value.CodedWOSform.ValueConcept
         
         Case NCID_PROVIDER_ROLE
            
            RoleNCID = oObsMod.ObsValue.value.CodedWOSform.ValueConcept
            
         Case NCID_SEMANTICLINKTORELATEDPROCEDURE
                     
            Set oProcedure = New EncProviderProcedure
            oProcedure.Init oObsMod.ObsValue.value.CodedWOSform.ValueConcept, ""
            oProcedure.State = rsSaved
            Procedures.Add oProcedure
            Set oProcedure = Nothing
                     
         End Select
         
      Next
      
   End With
   
   meCurrentState = rsSaved
   
   CDRLoad = True
   
ErrorHandler:
   Set oObsMod = Nothing
   Set oProcedure = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRSave"
   
End Function

Friend Function CDRSave(ByRef oEvent As ClinicalEvent) As Boolean

On Error GoTo ErrorHandler

Dim oClinObs      As ClinicalObservation
Dim oProcedure    As EncProviderProcedure

   If Val(NCID) > 0 And meCurrentState <> rsDelete Then
   
      'Create the clinical event for the provider
      Set oClinObs = AddClinObs(oEvent, NCID_ADDITIONALPROVIDER, NCID)
      
      'Create the clinical event for the provider
      If Val(RoleNCID) > 0 Then
         Call AddClinObsMod(oClinObs, NCID_PROVIDER_ROLE, RoleNCID)
      End If
      
      'Create the provider's semantic link for each of the procedures
      For Each oProcedure In Procedures
         If Val(oProcedure.DataId) Then
            If oProcedure.State = rsDelete Or oProcedure.State = rsRemove Then
               oProcedure.State = rsRemove
            Else
               Call AddClinObsMod(oClinObs, NCID_SEMANTICLINKTORELATEDPROCEDURE, oProcedure.DataId)      'SCR 51520 - Only save if not deleting
               oProcedure.State = rsSaved
            End If
         End If
      Next
            
      'Clean out procedures collection
      RemoveProcedures
         
   End If
   
   If meCurrentState = rsDelete Then
      meCurrentState = rsRemove
   Else
      meCurrentState = rsSaved
   End If
   
   CDRSave = True
   
ErrorHandler:
   Set oClinObs = Nothing
   Set oProcedure = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRSave"
   
End Function
Friend Function GEMS_SaveEncounterProvider(ByVal lEncounterNumber As Long, ByVal lFacilityNCID As Long, _
         ByVal lPriorityIndicator As Long) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim lIdx As Long

   Select Case meCurrentState
   
   Case rsNew
   
      sSql = "INSERT INTO Encounter_Provider " & _
            " ( EncounterNumber, FacilityNCID, Provider_NCID, Role_Type_NCID, Priority_Indicator) " & _
            " VALUES " & _
            " ( " & lEncounterNumber & ", " & lFacilityNCID & ", " & NCID & ", " & _
               RoleNCID & "," & lPriorityIndicator & " )"
               
   Case rsModified
      
      If msPrevNCID <> NCID And Procedures.Count > 0 Then
         Me.GEMS_DeleteProviderProcedures CStr(lEncounterNumber), CStr(lFacilityNCID), False
      End If
   
      sSql = "UPDATE Encounter_Provider " & _
            " SET Provider_NCID = " & NCID & ", " & _
               " Role_Type_NCID = " & RoleNCID & ", " & _
               " Priority_Indicator = " & lPriorityIndicator & _
            " WHERE EncounterNumber = " & lEncounterNumber & _
               " AND FacilityNCID = " & lFacilityNCID & _
               " AND Provider_NCID = " & msPrevNCID
               
   Case rsDelete
   
      'Remove all procedures first
      Me.GEMS_DeleteProviderProcedures CStr(lEncounterNumber), CStr(lFacilityNCID), True
      
      sSql = "DELETE FROM Encounter_Provider " & _
            " WHERE EncounterNumber = " & lEncounterNumber & _
            "     AND FacilityNCID = " & lFacilityNCID & _
            "     AND Provider_NCID = " & msPrevNCID
            
   End Select
   
   If Len(sSql) > 0 Then
      If oSQL Is Nothing Then Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
      oSQL.Execute sSql
      
      If Me.State = rsDelete Then
         Me.State = rsRemove
      Else
         Me.State = rsSaved
      End If
      msPrevNCID = NCID
      
   End If
   
   GEMS_SaveEncounterProvider = True
   
ErrorHandler:
   Set oSQL = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_SaveEncounterProvider"
      
End Function

Friend Function GEMS_SaveProviderProcedures(ByVal sEncounterNumber As String, ByVal sFacilityNCID As String) As Boolean

On Error GoTo ErrorHandler

Dim oProcedure As EncProviderProcedure
Dim bSuccess As Boolean


   If mcProcedures Is Nothing Then
      GEMS_SaveProviderProcedures = True
   Else
   
      'Save Each Procedure
      bSuccess = True
      For Each oProcedure In mcProcedures
         bSuccess = bSuccess And oProcedure.GEMS_SaveProcedure(sEncounterNumber, sFacilityNCID, NCID)
      Next
      
      GEMS_SaveProviderProcedures = bSuccess
      
      RemoveProcedures
      
   End If
   
ErrorHandler:
   Set oProcedure = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_SaveProviderProcedures"

End Function

Public Property Get IsDirty() As Boolean

On Error GoTo ErrorHandler

Dim oProc As EncProviderProcedure

   If State = rsNew Or State = rsDelete Or State = rsModified Then
      IsDirty = True
   Else
      'Check the procedures
      For Each oProc In Procedures
         If oProc.IsDirty Then
            IsDirty = True
            Exit For
         End If
      Next
   End If
   
ErrorHandler:
   Set oProc = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "IsDirty"
   
End Property

Private Sub RemoveProcedures()

Dim lIdx As Long
Dim oProcedure As EncProviderProcedure

      If mcProcedures Is Nothing Then Exit Sub
      
      'Clean out the procedures collection
      For lIdx = mcProcedures.Count To 1 Step -1
         Set oProcedure = mcProcedures(lIdx)
         If oProcedure.State = rsRemove Then
            mcProcedures.Remove (lIdx)
         End If
      Next
      
End Sub

Public Function ADMRole() As String
   Select Case CLng(Trim$(msRoleNCID))
      Case rtnAttendingProvider:    ADMRole = "1"
      Case rtnAssistingProvider:    ADMRole = "2"
      Case rtnSupervisingProvider:  ADMRole = "3"
      Case rtnNurse:                ADMRole = "4"
      Case rtnParaProfessional:     ADMRole = "5"
      Case rtnOperatingProvider1:   ADMRole = "6"
      Case Else:                    ADMRole = vbNullString
   End Select
End Function

Public Property Get State() As enumRecordState
  State = meCurrentState
End Property

Public Property Let State(ByVal eNewValue As enumRecordState)
  meCurrentState = eNewValue
End Property

Public Property Get PrevNCID() As String:                    PrevNCID = msNCID:                  End Property
Public Property Let PrevNCID(ByVal sData As String):         msPrevNCID = sData:                 End Property
Public Property Get Procedures() As Collection
   If mcProcedures Is Nothing Then
      Set mcProcedures = New Collection
   End If
   Set Procedures = mcProcedures
End Property

Public Function GEMS_DeleteProviderProcedures(ByVal sEncounterNumber As String, ByVal sFacilityNCID As String, _
      ByVal bRemoveAll As Boolean) As Boolean

On Error GoTo ErrorHandler

Dim lProcIdx As Long
Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps

   'Remove any procedures and reset to new to avoid constraint violations
   sSql = "DELETE FROM Encounter_Procedure " & _
          " WHERE EncounterNumber = " & sEncounterNumber & _
         " AND FacilityNCID = " & sFacilityNCID & _
         " AND Provider_NCID = " & msPrevNCID
   Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
   oSQL.Execute sSql
   
   'Reset any procedures that would need to be saved
   For lProcIdx = 1 To Procedures.Count
      If bRemoveAll Then
         Procedures(lProcIdx).State = rsRemove
      Else
         'Set state based upon appropriate action
         Select Case Procedures(lProcIdx).State
         Case rsModified, rsSaved
            Procedures(lProcIdx).State = rsNew
         Case rsDelete
            Procedures(lProcIdx).State = rsRemove
         End Select
      End If
   Next

   'Clean out procedures collection
   RemoveProcedures
   
   GEMS_DeleteProviderProcedures = True
   
ErrorHandler:
   Set oSQL = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_DeleteProviderProcedures"
                  
End Function

