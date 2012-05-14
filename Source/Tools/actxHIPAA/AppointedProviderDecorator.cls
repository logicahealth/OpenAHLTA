VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppointedProviderDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE_NAME As String = "AppointedProviderDecorator"

Private mlCtrlCount As Long
Private mcClinicProviders As Collection
Private mcControls As Collection

Public Sub AddEncounterProcedure(ByVal sDataId As String, ByVal sSnoID As String, ByVal sDescription As String)

On Error GoTo Error_Handler

Dim lCntrlIdx As Long

   For lCntrlIdx = 1 To mlCtrlCount
      mcControls(lCntrlIdx).AddEncounterProcedure sDataId, sSnoID, sDescription
   Next
   
Error_Handler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "AddEncounterProcedure"
   
End Sub
 
Public Function Validate(ByRef sWarning As String) As Boolean

On Error GoTo Error_Handler

Dim lIdxInner As Long
Dim lIdxOuter As Long
   
   Validate = True
   
   'Check for duplicate provider
   For lIdxOuter = 1 To mcControls.Count
      For lIdxInner = lIdxOuter + 1 To mcControls.Count
         If mcControls(lIdxOuter).ProviderNCID = mcControls(lIdxInner).ProviderNCID And _
            mcControls(lIdxOuter).ProviderNCID <> vbNullString And mcControls(lIdxOuter).ProviderNCID <> "0" Then
               sWarning = "Providers may not be associated with an encounter more than once!" & vbCrLf
               Validate = False
               Exit Function
         End If
      Next
   Next
   
   'Make sure all providers have a role associated with them
   For lIdxOuter = 1 To mcControls.Count
      'SCR 48996 - Provider role can also be equal to a "0" when its missing
      If mcControls(lIdxOuter).RoleNCID = 0 And mcControls(lIdxOuter).ProviderNCID <> vbNullString And mcControls(lIdxOuter).ProviderNCID <> "0" Then
         sWarning = "A Role is required for each Provider." & vbCrLf
         Validate = False
         Exit Function
      End If
   Next
   
Error_Handler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "Validate"
      
End Function

Public Function DebugPrint() As Boolean

On Error GoTo Error_Handler

Dim bSuccess As Boolean
Dim oCntrl As AppointedProviderSelector
   
   bSuccess = True
      
   'Collect the data for each control and save the data
   For Each oCntrl In mcControls
      bSuccess = bSuccess And oCntrl.DebugPrint()
   Next
   
   DebugPrint = bSuccess
   
Error_Handler:
   'Clean Up Objects
   Set oCntrl = Nothing
   
   If CBool(Err) Then GlobalShared.ShowVBError Err.Number, Err.Description, "AppointedProviderDecorator.DebugPrint", COMPONENT_ERR_TITLE, vbExclamation
      
End Function

Public Sub Init(ByRef oLogon As Logon, ByVal bAPV As Boolean, ByVal sClinicNCID As String, ParamArray aSelectorControls())

On Error GoTo Error_Handler

Dim vParameter As Variant
Dim oAppointedProvider As AppointedProviderSelector
Dim oClinicOps As ClinicOps
   
   mlCtrlCount = 0
   
   Set mcControls = New Collection
   
   'Load Clinic Providers - SCR 51226: Retreive all providers not just those with signing rights
   Set oClinicOps = New ClinicOps
   Set mcClinicProviders = oClinicOps.GetClinicProviders(sClinicNCID, False)
        
   For Each vParameter In aSelectorControls
   
      'Make sure we don't have any invalid controls in parameter array
      If TypeOf vParameter Is AppointedProviderSelector Then
         
         'Initialize Each Control
         Set oAppointedProvider = vParameter
         Set oAppointedProvider.ProviderCollection = mcClinicProviders
         oAppointedProvider.InitNew oLogon, bAPV, sClinicNCID
         mcControls.Add oAppointedProvider
         
         'Initialize additional properties used only by the decorator control
         Set oAppointedProvider.ControlDecorator = Me
         oAppointedProvider.RoleCaption = "Role"
         If mlCtrlCount = 0 Then
            'The first control defaults to the current provider and cannot be modified
            oAppointedProvider.ProviderSelectionEnabled = False
            'SCR 43277 and 43279 - Before setting the provider, ensure the provider exists in the provider collection by attempting to add it - duplicates will not be added
            AddNewProvider oLogon.UserName, oLogon.UserNCID
            oAppointedProvider.ProviderNCID = oLogon.UserNCID
            oAppointedProvider.ProviderCaption = "Appointed Provider"
         Else
            oAppointedProvider.ProviderSelectionEnabled = True
            oAppointedProvider.ProviderCaption = "Additional Provider #" & CStr(mlCtrlCount)
         End If
            
         'Add control reference to interal control collection and increment control count
         mlCtrlCount = mlCtrlCount + 1
         
      End If
   Next
   
Error_Handler:
   'Clean Up Objects
   Set oAppointedProvider = Nothing
   Set oClinicOps = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "Init"
      
End Sub

'Provides a method of notification so that a new provider that has been added to the provider collection can be added to all
'of the controls that are under the management of the decorator class
Friend Sub AddNewProvider(ByVal sNewProviderName As String, ByVal sNewProviderNCID As String)

On Error GoTo Error_Handler

Dim oAppointedProviderSelector As AppointedProviderSelector
Dim oNewProvider As Provider

   'Check for existance of provider in the collection and exit if they are already there
   If Not GetCollectionItem(mcClinicProviders, sNewProviderNCID) Is Nothing Then Exit Sub
   
   'Create the new provider object and add to the providers collection
   Set oNewProvider = New Provider
   oNewProvider.ProviderName = sNewProviderName
   oNewProvider.ClinicianNCID = sNewProviderNCID
   mcClinicProviders.Add oNewProvider, oNewProvider.ClinicianNCID
   
   'Now add the item to each of the controls provider list
   For Each oAppointedProviderSelector In mcControls
      oAppointedProviderSelector.AddProviderItem oNewProvider.ProviderName, oNewProvider.ClinicianNCID
   Next
   
Error_Handler:
   'Clean Up Objects
   Set oAppointedProviderSelector = Nothing
   Set oNewProvider = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "AddNewProvider"
   
End Sub

Public Sub SetProviderInfo(ByVal lCntrlIdx As Long, ByVal sProviderNCID As String, ByVal sRoleNCID As String, ByVal sProviderName As String)

On Error GoTo Error_Handler

   'Validate the control index
   If lCntrlIdx <= mcControls.Count Then
      mcControls(lCntrlIdx).SetProviderInfo sProviderNCID, sRoleNCID, sProviderName
   End If

Error_Handler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "SetProviderInfo"
   
End Sub
