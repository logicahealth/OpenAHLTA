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
Attribute VB_Name = "ClinicOpsEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const msMODULE_NAME         As String = "ClinicOpsEx"

Private Enum penmColType
   enmProvider
   enmClinic
End Enum

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''
Private mMapperCC As DL_Interfaces.iMapperClinicClient

Public mbUseClinicProviderTable     As Boolean  'Default to True

Public Function GetClinic(sNCID As String, Optional sFacilityNCID As String = vbNullString) As DTOs.Clinic
  Const sPROC_NAME  As String = "GetClinic"
  Dim ocolClinics As DTOs.Clinics
  Dim Params As DM_ClinicClient.ClinicSearchParam
   
On Error GoTo errHandler

    Set Params = New DM_ClinicClient.ClinicSearchParam
    Params.ClinicNCID = sNCID
    Params.FacilityNCID = sFacilityNCID
    Params.UseClinicProviderTable = mbUseClinicProviderTable
    Set ocolClinics = mMapperCC.GetClinics(Params)
    
    If ocolClinics.Count > 0 Then
      Set GetClinic = ocolClinics(1)
    Else
      Set GetClinic = Nothing
    End If
Exit Function
errHandler:
  Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
  Set GetClinic = New DTOs.Clinic
End Function

Public Function GetAllClinics(rsMTFFacilityNCID As String, Optional boolIncludeNonMTF As Boolean = True) As DTOs.Clinics
  Const sPROC_NAME  As String = "GetAllClinics"
  Dim Params As DM_ClinicClient.ClinicSearchParam
    
On Error GoTo errHandler

    Set Params = New DM_ClinicClient.ClinicSearchParam
    Params.FacilityNCID = rsMTFFacilityNCID
    Params.IncludeNonMTF = boolIncludeNonMTF
    Params.UseClinicProviderTable = mbUseClinicProviderTable
    
    Set GetAllClinics = mMapperCC.GetClinics(Params)
   
Exit Function
errHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   Set GetAllClinics = New DTOs.Clinics
   
End Function

Public Function GetClinicProvidersByRole _
      (ByVal vsClinicNCID As String, Optional ByVal sRole As String = "") As Collection
Const sPROC_NAME  As String = "GetClinicProvidersByRole"
      
On Error GoTo errHandler

'***** Need to keep for now.  Will be replaced when we can get Security Object User Roles
   Set GetClinicProvidersByRole = ClinicClient_colGetClinicProviders(vsClinicNCID, True, sRole, True)
Exit Function

errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
    Set GetClinicProvidersByRole = New Collection
End Function

Public Function GetClinicsWorkload _
      (ByVal sMTFFacilityNCID As String, ByRef colClinics As DTOs.Clinics) As Boolean
  Const sPROC_NAME     As String = "GetClinicsWorkload"

  Dim Params As DM_ClinicClient.ClinicSearchParam
   
On Error GoTo errHandler

  Set Params = New DM_ClinicClient.ClinicSearchParam
  Params.FacilityNCID = sMTFFacilityNCID
  Params.IncludeWorkload = True
  Params.UseClinicProviderTable = mbUseClinicProviderTable
  
  GetClinicsWorkload = mMapperCC.GetClinicsWorkload(Params, colClinics)
   
Exit Function
errHandler:
   GetClinicsWorkload = False
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)

End Function

Public Function GetClinicWorkloadType _
      (ByVal sFacilityNCID, ByVal sClinicNCID As String, ByVal sProviderNCID As String, ByVal sAppointmentCode As String, _
       Optional ByRef sErrMsg As String, Optional ByVal bPopMsgOnErr As Boolean = False _
      ) As String
  Const sPROC_NAME     As String = "GetClinicWorkloadType"
    
  Dim ocolClinics As DTOs.Clinics
  Dim Params As DM_ClinicClient.ClinicSearchParam
  Dim sResult As String
  
On Error GoTo errHandler

  Set Params = New DM_ClinicClient.ClinicSearchParam
  Params.FacilityNCID = sFacilityNCID
  Params.ClinicNCID = sClinicNCID
  Params.UserNCID = sProviderNCID
  Params.ApptCode = sAppointmentCode
  Params.IncludeWorkload = True
  Params.UseClinicProviderTable = mbUseClinicProviderTable
  
  Set ocolClinics = mMapperCC.GetClinics(Params)
  
  If ocolClinics.Count > 0 Then
    GetClinicWorkloadType = ocolClinics.Item(1).Default_Workload_Type
  Else
    GetClinicWorkloadType = ""
  End If
    
Exit Function
errHandler:
   GetClinicWorkloadType = ""
   sErrMsg = Err.Description
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetProviderExt(ByVal sProviderNCID As String) As DTOs.Providers
   Const sPROC_NAME  As String = "GetProviderExt"
    Dim ocolProviders As DTOs.Providers
    Dim Params As DM_ClinicClient.ProviderSearchParam
   
On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.ProviderNCID = sProviderNCID
   Params.IncludeExtendedData = True
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetProviderExt = mMapperCC.GetProviders(Params)
   
Exit Function
errHandler:
   Set GetProviderExt = New DTOs.Providers
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
'   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sProviderNCID)

End Function

Public Function GetProvidersWorkload _
   (ByVal sMTFFacilityNCID As String, ByVal sClinicNCID As String, ByRef colProviders As DTOs.Providers) _
    As Boolean
   Const sPROC_NAME     As String = "GetProvidersWorkload"
    
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler

  Set Params = New DM_ClinicClient.ProviderSearchParam
  Params.FacilityNCID = sMTFFacilityNCID
  Params.ClinicNCID = sClinicNCID
  Params.UseClinicProviderTable = mbUseClinicProviderTable
  
  GetProvidersWorkload = mMapperCC.GetProvidersWorkload(Params, colProviders)
 
Exit Function
errHandler:
   GetProvidersWorkload = False
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & gsLINE & Erl, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function GetUserClinics(ByVal vsMTFFacilityNCID As String, ByVal vsUserNCID As String, Optional ByVal vbGetIBWAClinics As Boolean = False) As DTOs.Clinics
   Const sPROC_NAME  As String = "GetUserClinics"
   
   Dim ocolProviders As DTOs.Providers
   Dim Params As DM_ClinicClient.ClinicSearchParam
   
On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ClinicSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.UserNCID = vsUserNCID
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   Params.IncludeIBWAClinics = vbGetIBWAClinics
   
   Set GetUserClinics = mMapperCC.GetClinics(Params)

Exit Function
errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & " usererNcid=" & vsUserNCID, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vsUserNCID)
   Set GetUserClinics = New DTOs.Clinics
   
End Function

Public Function GetProviderName _
      (ByVal vsMTFFacilityNCID As String, ByVal vsProviderNCID As String) As String
   
   Const sPROC_NAME  As String = "GetProviderName"
   Dim ocolProviders As DTOs.Providers
   Dim Params As DM_ClinicClient.ProviderSearchParam
On Error GoTo errHandler

   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.ProviderNCID = vsProviderNCID
   
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set ocolProviders = mMapperCC.GetProviders(Params)
   
   If ocolProviders.Count > 0 Then
    GetProviderName = ocolProviders(1).Name
   Else
    GetProviderName = ""
   End If
Exit Function

errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetProviderClinics _
      (ByVal vsMTFFacilityNCID As String, ByVal vsClinicianNCID As String) As DTOs.Clinics
   Const sPROC_NAME  As String = "GetProviderClinics"
   
   Dim Params As DM_ClinicClient.ClinicSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ClinicSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.UserNCID = vsClinicianNCID
   Params.UseClinicProviderTable = mbUseClinicProviderTable
    
   Set GetProviderClinics = mMapperCC.GetClinics(Params)
Exit Function
errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & " clinicianNCID=" & vsClinicianNCID, sPROC_NAME, vbExclamation)
   Set GetProviderClinics = New DTOs.Clinics
   
End Function

Public Function GetClinicProviders _
      (ByVal vsClinicNCID As String, Optional ByVal bUsersOnly As Boolean = False, _
       Optional ByVal bIncludeRoles As Boolean _
      ) As DTOs.Providers
   Const sPROC_NAME  As String = "GetClinicProviders"
   
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.ClinicNCID = vsClinicNCID
   Params.ClinicUsersOnly = bUsersOnly
   Params.IncludeRoles = bIncludeRoles
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetClinicProviders = mMapperCC.GetProviders(Params)
   
Exit Function
errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetClinicProvidersLikeLastName _
      (ByVal vsMTFFacilityNCID, _
      ByVal vsLikeLastName, _
      ByVal vsClinicNCID As String, _
      Optional ByVal bUsersOnly As Boolean = False) As DTOs.Providers
   Const sPROC_NAME  As String = "GetClinicProvidersLikeLastName"
   
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.ProviderLastName = vsLikeLastName
   Params.ClinicNCID = vsClinicNCID
   Params.ClinicUsersOnly = bUsersOnly
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetClinicProvidersLikeLastName = mMapperCC.GetProviders(Params)
   
Exit Function
errHandler:
   Set GetClinicProvidersLikeLastName = New DTOs.Providers
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   
End Function

Public Function GetProviders _
   (ByVal sMTFFacilityNCID As String, _
   ByVal sLikeLastName As String, _
   ByVal sClinicNCID As String, _
   Optional ByVal bUsersOnly As Boolean = False, _
   Optional ByVal bSignable As Boolean = False) As DTOs.Providers
   
   Const sPROC_NAME  As String = "GetProviders"
   
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = sMTFFacilityNCID
   Params.ProviderLastName = sLikeLastName
   Params.ClinicNCID = sClinicNCID
   Params.ClinicUsersOnly = bUsersOnly
   Params.SignableOnly = bSignable
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetProviders = mMapperCC.GetProviders(Params)

Exit Function
errHandler:
   Set GetProviders = New DTOs.Providers
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function
Public Function GetClinicIEN(ByVal vsClinicNCID As String) As String
  Const sPROC_NAME  As String = "GetClinicIEN"
  
  Dim ocolClinics As DTOs.Clinics
  Dim Params As DM_ClinicClient.ClinicSearchParam

On Error GoTo errHandler

  Set Params = New DM_ClinicClient.ClinicSearchParam
  Params.ClinicNCID = vsClinicNCID
  Params.UseClinicProviderTable = mbUseClinicProviderTable
  
  Set ocolClinics = mMapperCC.GetClinics(Params)
  
  If ocolClinics.Count > 0 Then
    GetClinicIEN = ocolClinics(1).IEN
  Else
    GetClinicIEN = ""
  End If
Exit Function

errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetClinicNCID(ByVal vsMTFFacilityNCID As String, vsClinicIEN As String) As String
  Const sPROC_NAME  As String = "GetClinicNCID"
  
  Dim ocolClinics As DTOs.Clinics
  Dim Params As DM_ClinicClient.ClinicSearchParam
  
On Error GoTo errHandler

  Set Params = New DM_ClinicClient.ClinicSearchParam
  Params.FacilityNCID = vsMTFFacilityNCID
  Params.ClinicIEN = vsClinicIEN
  Params.UseClinicProviderTable = mbUseClinicProviderTable
  
  Set ocolClinics = mMapperCC.GetClinics(Params)
  
  If ocolClinics.Count > 0 Then
    GetClinicNCID = ocolClinics(1).NCID
  Else
    GetClinicNCID = ""
  End If
Exit Function

errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetAllProviders _
   (ByVal vsMTFFacilityNCID, _
    Optional ByVal bUsersOnly As Boolean = False, Optional ByVal bSignable As Boolean = False, _
    Optional ByVal sSingleProviderNCID As String = vbNullString) As DTOs.Providers
    
   Const sPROC_NAME     As String = "GetAllProviders"

   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.ClinicUsersOnly = bUsersOnly
   Params.SignableOnly = bSignable
   Params.ProviderNCID = sSingleProviderNCID
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetAllProviders = mMapperCC.GetProviders(Params)
   
Exit Function
errHandler:
   Set GetAllProviders = New DTOs.Providers
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
    
End Function

Public Function GetAllProvidersLikeLastName _
   (ByVal vsMTFFacilityNCID, ByVal vsLikeLastName, _
    Optional ByVal bUsersOnly As Boolean = False, Optional ByVal bSignable As Boolean = False _
    ) As DTOs.Providers
    
   Const sPROC_NAME     As String = "GetAllProvidersLikeLastName"
   
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.ProviderLastName = vsLikeLastName
   Params.ClinicUsersOnly = bUsersOnly
   Params.SignableOnly = bSignable
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetAllProvidersLikeLastName = mMapperCC.GetProviders(Params)
   
Exit Function
errHandler:
   Set GetAllProvidersLikeLastName = New DTOs.Providers
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetProviderIEN(ByVal vsMTFFacilityNCID As String, ByVal vsProviderNCID As String) As String
   Const sPROC_NAME     As String = "GetProvidersIEN"
   
   Dim ocolProviders As DTOs.Providers
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler

   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.ProviderNCID = vsProviderNCID
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set ocolProviders = mMapperCC.GetProviders(Params)
   
   If ocolProviders.Count > 0 Then
     GetProviderIEN = ocolProviders(1).IEN
   Else
     GetProviderIEN = ""
   End If
Exit Function
errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Public Function GetProviderNCID(ByVal vsMTFFacilityNCID As String, ByVal vsProviderIEN As String) As String
   Const sPROC_NAME     As String = "GetProviderNCID"
   
   Dim ocolProviders As DTOs.Providers
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler

   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.FacilityNCID = vsMTFFacilityNCID
   Params.ProviderIEN = vsProviderIEN
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set ocolProviders = mMapperCC.GetProviders(Params)
   
   If ocolProviders.Count > 0 Then
     GetProviderNCID = ocolProviders(1).NCID
   Else
     GetProviderNCID = ""
   End If
Exit Function
errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function

Private Sub Class_Initialize()
  Set mMapperCC = gobjShared.DataAccessService.ClinicClientMapper
   
  mbUseClinicProviderTable = True  '''<bjr102701 SCR4004
  
End Sub

Public Function GetSignableClinicProviders _
   (ByVal vsClinicNCID As String, Optional ByVal bUsersOnly As Boolean = False) As DTOs.Providers
   
   Const sPROC_NAME  As String = "GetSignableClinicProviders"
   
   Dim Params As DM_ClinicClient.ProviderSearchParam

On Error GoTo errHandler
   Set Params = New DM_ClinicClient.ProviderSearchParam
   Params.ClinicNCID = vsClinicNCID
   Params.ClinicUsersOnly = bUsersOnly
   Params.SignableOnly = True
   Params.UseClinicProviderTable = mbUseClinicProviderTable
   
   Set GetSignableClinicProviders = mMapperCC.GetProviders(Params)
   
Exit Function
errHandler:
   Set GetSignableClinicProviders = New DTOs.Providers
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
End Function
