VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListManagerOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MODULE_NAME As String = "CodeDomainOps"

Public Type GeographicLocation
   glIEN As Long
   gsDesc As String
End Type

Public Enum CodeDesignationTypes
   cdtRounds = 1
   cdtAccidentInjury
   cdtPregnancy
   cdtInvalidRNDs
End Enum

Public Enum CodeSetTypes
   cstEandM = 1
   cstDiagnosis
   cstProcedure
End Enum

'This method will be used by A/P to determine when a diagnosis or procedure is related to pregnancy and display the pregnancy dialog.  Parameter sType will be either 'P' for procedure or 'D' for diagnosis.
Public Function IsPregnancyCode(ByVal eCodeType As CodeSetTypes, ByVal sCode As String) As Boolean
   IsPregnancyCode = gcCodeDesignations.IsCodeInDomain(cdtPregnancy, eCodeType, sCode)
End Function

Public Function IsRoundsCode(ByVal eCodeSetType As CodeSetTypes, ByVal sCode As String) As Boolean
   IsRoundsCode = gcCodeDesignations.IsCodeInDomain(cdtRounds, eCodeSetType, sCode)
End Function

'This method will be used by A/P to determine when to display accident injury selection dialog.
Public Function IsAccidentCode(ByVal eCodeSetType As CodeSetTypes, ByVal sCode As String) As Boolean
   IsAccidentCode = gcCodeDesignations.IsCodeInDomain(cdtAccidentInjury, eCodeSetType, sCode)
End Function

'Returns a collection of locations that exist at the facility
Public Function GetFacilityLocations(ByVal lFacilityNCID As Long) As Locations
   If gcFacilities Is Nothing Then
      Set gcFacilities = New Facilities
   End If
   Set GetFacilityLocations = gcFacilities.Locations(lFacilityNCID)
End Function

Private Sub Class_Initialize()
   If gcCodeDesignations Is Nothing Then
      Set gcCodeDesignations = New CodeDesignations
   End If
End Sub

Public Function IsCodeInDomain(ByVal eCodeDesignation As CodeDesignationTypes, _
         ByVal eCodeSetType As CodeSetTypes, ByVal sCode As String) As Boolean
   IsCodeInDomain = gcCodeDesignations.IsCodeInDomain(eCodeDesignation, eCodeSetType, sCode)
End Function

