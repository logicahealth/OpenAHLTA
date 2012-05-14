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
Attribute VB_Name = "cCISM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Private myCISMPatients As Patients
'Private myAllergies As Allergies
'Private myImmunizations As Immunizations

Public Sub GetAllergies(ByVal sFMP As String, ByVal sSSN As String, colAllergies As Object)
Dim lIEN As Long
'Dim myPatient As mmmhispatient.Patient
'Dim myCISMPatient As Patient
'Set myCISMPatient = New CISMTypes.Patient

Set colAllergies = New Allergies

lIEN = GetUnitNumber(sFMP, sSSN)
If lIEN = 0 Then
  Exit Sub
End If

GetAllergiesFromCDR lIEN, sFMP, sSSN, colAllergies


End Sub
Public Sub GetProviders(ByRef ProviderName() As String, ByRef ProviderNCID() As String)
On Error GoTo Cleanup
'modCISM.GetProvidersFromCDR ProviderName, ProviderNCID
'objXlate.GetProviders ProviderName, ProviderNCID


Exit Sub
Cleanup:
WriteMessage "blew up in class module." & vbCrLf & Err.Number & ": " & Err.Description, "cCISM.GetProviders"

    
End Sub
Public Function GetPatientInfo(ByRef objPatient As Object, ByRef colAllergies As Object, ByRef colImmunizations As Object) As Boolean
Dim myPatients As New Patients
Dim ThisPatient As Patient
Dim myAllergies As New Allergies
Dim ThisAllergy As Allergy
Dim myImmunizations As New Immunizations
Dim ThisImmunization As Immunization
Dim ThatPatient As MMMHISPatient.Patient
Dim myssn As String
Dim myfmp As String
Dim myIEN As Long

On Error GoTo handler
  
  GetPatientInfo = False
  'working code goes here
  Set ThisPatient = objPatient
  myIEN = GetUnitNumber(ThisPatient.sFMP, ThisPatient.sSSN)
  ThisPatient.lIEN = myIEN
  Set ThatPatient = getPatient(myIEN)
  ThisPatient.dDOB = ThatPatient.DateOfBirth
  ThisPatient.sFirstName = ThatPatient.Name
  ThisPatient.sLastName = ThatPatient.Name
  GetAllergiesFromCDR myIEN, ThisPatient.sFMP, ThisPatient.sSSN, myAllergies
  Call GetImmunizationsFromCDR(myImmunizations, ThisPatient.sSSN, ThisPatient.sFMP, myIEN)
  
  'working code ends here
  
  GetPatientInfo = True
  Set objPatient = ThisPatient
  Set colAllergies = myAllergies
  Set colImmunizations = myImmunizations

  
Exit Function
handler:
  WriteMessage Erl & " " & Err.Number & ": " & Err.Description, "GetPatientInfo"
  GetPatientInfo = False
  Set objPatient = Nothing
  Set colAllergies = Nothing
  Set colImmunizations = Nothing
  Set ThatPatient = Nothing
  Set ThisAllergy = Nothing
  Set ThisPatient = Nothing
  Set ThisImmunization = Nothing
  GetPatientInfo = False
End Function

Public Function GetVaccines(Name() As String, NCID() As String)
  objXlate.GetVaccines Name, NCID
End Function

Public Function GetAllergens(Name() As String, TypeNCIDs() As String, ncids() As String)
On Error GoTo Cleanup
Dim objconfig As MMMHISConfiguration.Config
   
1 Set objconfig = New MMMHISConfiguration.Config
2 If Not objconfig.RetrieveConfiguration(gobjLogon.EnterpriseNCID, _
                                       gobjLogon.FacilityNCID, _
                                       gobjLogon.RoleNCID, _
                                       gobjLogon.UserNCID) Then
3  WriteMessage "Unable to retrieve 3MCareInnovation configuration information.", "cCISM.GetAllergens"
  GoTo Cleanup
End If
4 objXlate.GetAllergens Name, TypeNCIDs, ncids, objconfig


Set GetAllergens = Nothing
Set objconfig = Nothing
Exit Function
Cleanup:
  WriteMessage "Line " & Erl & ": " & Err.Number & ": " & Err.Description, "cCISM.GetAllergens"
  Set objconfig = Nothing
  
End Function

Public Sub SaveImmunizations(colImmunizations As Object)
Dim mcolImmunizations As Object
Dim myImmunization As Object
Dim strTemp As String
WriteMessage "in save imm.", "ccism.saveimmunizations"
On Error GoTo Cleanup
 If objXlate Is Nothing Then
   Set objXlate = New XlateNCID
   objXlate.Setup gobjLogon.EnterpriseNCID
 End If
 Set mcolImmunizations = colImmunizations

 'for each patients' immunizations:
 For Each myImmunization In mcolImmunizations
 With myImmunization
  If Len(.sSSN) <> 9 Then
    GoTo EndOfImmunization
  End If
   .lIEN = GetUnitNumber(.sFMP, .sSSN)

  If .lIEN = 0 Then
    GoTo EndOfImmunization
  End If
WriteMessage "got IEN: " & .lIEN, "ccism.saveimmunizations"
  .sSiteNCID = objXlate.SiteNCID(.sSiteName) '"33963"
  .sAdminLocationName = gobjLogon.FacilityName
  .sAdminLocationNCID = gobjLogon.FacilityNCID
  .sClinicianNCID = gobjLogon.StoringUserNCID '
  .sClinicianNCID = objXlate.ClinicianNCID(.sClinicianName) '
  .sManufacturerNCID = objXlate.ManufacturerNCID(.sManufacturerName)  '"206734"
  .sRouteNCID = objXlate.RouteNCID(.sRouteName) '"67590"
  .sUnitsNCID = objXlate.UnitsNCID(.sUnitsName)
  .sVaccineName = objXlate.VaccineName(.sVaccineNCID)
  .sSeqNumberNCID = objXlate.SeqNumberNCID(.sSeqNumberName)

WriteMessage "converted ncids", "ccism.saveimmunizations"

EndOfImmunization:

'WriteMessage "completed conversions", "ccism.saveimmunizations"
End With

  If myImmunization.lIEN <> 0 Then
     If myImmunization.sEditFlag = "D" Then
       If myImmunization.sVaccineNCID = "107572" Then
          RemoveTestFromCDR myImmunization
       Else
          RemoveImmunizationFromCDR myImmunization
       End If
     Else
       If myImmunization.sVaccineNCID = "107572" Then  'tb test
         WriteTestToCDR myImmunization
       Else
'WriteMessage "about to call writeimm.", "ccism.saveimmunizations"
         WriteImmunizationsToCDR myImmunization.lIEN, myImmunization
       End If
     End If
  End If
Next

'Set myCISMPatient = Nothing ' As Object ' CISM.Patient
'Set my3MPatient = Nothing 'As MMMHISPatient.Patient
Set mcolImmunizations = Nothing 'As Object 'Immunizations
Set myImmunization = Nothing 'Immunization

Exit Sub
Cleanup:
WriteMessage "Line " & Erl & ": " & Err.Number & ": " & Err.Description, "cCISM.SaveImmunizations"
'Resume Next
Set mcolImmunizations = Nothing
Set myImmunization = Nothing

End Sub



Public Function getPatientIEN(SSN As String) As Long

getPatientIEN = GetUnitNumber("20", SSN)

End Function

Private Sub Class_Initialize()
    WriteMessage "Initializing CISM Class", "CISM"
    'Call DoSetup("")
End Sub

Private Sub Class_Terminate()
    WriteMessage "CISM Class Terminate", "CISM"
  'Set myCISMPatients = Nothing
  'Set myAllergies = Nothing
  'Set myImmunizations = Nothing
  'DoShutdown
End Sub
