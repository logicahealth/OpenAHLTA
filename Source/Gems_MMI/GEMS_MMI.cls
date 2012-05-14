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
Attribute VB_Name = "GEMS_MMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarUnitNumber              As Long
Private mvarBirthDate               As Date
Private mvarMaritalStatus           As String '- coded
Private mvarNCID                    As String
Private mvarPCM                     As String
Private mvarPCMAddress              As String
Private mobjPCMDutyPhone            As GEMS_Phone
Private mvarPCMLocation             As String
Private mvarPCMNCID                 As String
Private mvarPCMSpecialty            As String
Private mvarRace                    As String '- coded
Private mvarReligion                As String '- coded
Private mvarSEX                     As String
Private mvarSSN                     As String
Private mvarFacilityNCID            As String
Private mvarEncounterNumber         As String
Private mvarEncounterFacilityNCID   As String
Private mvarMilitaryStatus          As String '- not yet available
Private mvarMedRecordNumber         As String
Private mobjAddresses               As GEMS_Addresses
Private mobjNames                   As Gems_Names    'added WR 5/13/2002
Private mobjPhones                  As GEMS_Phones
Private mobjGenericIDs              As GEMS_Generic_IDs
Private mobjGenericDatas            As GEMS_Generic_Datas
'this will go away in future version WR
Private mobjName                    As GEMS_Name
'*-- Coleman 052501: Support of MMI_Site in Gems
Private mobjSites                   As GEMS_Sites
Private mobjSponsor                 As GEMS_MMI
Private mcolCommandInterest         As Collection    ' of strings
Private mcolCommandSecurity         As Collection    ' of strings
'*-- Coleman 100201
Private mcolSpecialWorkStatus       As Collection    ' of strings
'--GEMS added for field use
Private mvarDeployLoc               As String
Private mvarDeployUnit              As String
Private mvarDNBI                    As String

Private mobjHomephone               As GEMS_Phone
Private mvarFullName                As String
Private mvarUpdateFlag              As String
Private mvarNextGDOrdinal           As Long      'the next available ordinal number
Private mbForeignID                 As Boolean
'Private msForeignID                 As String

'*-- Coleman 061301: Four properties stored in Generic_IDs
'Private mvarFMPSponsorSSN           As String
'Private mvarSponsorSrchString       As String
'Private mvarPatientSrchString       As String
'Private msSponsorSSN                As String

'*-- Coleman 051601: Six patient "Flag" properties added
Private mvarAdoptionFlag            As String
Private mvarExpiredFlag             As String
Private mvarGuarantorFlag           As String
Private mvarOraganDonorFlag         As String
Private mvarOrganDonorFlag          As String
Private mvarPatientFlag             As String
Private mvarUnidentifiedFlag        As String

'*-- Coleman 062101: Added two properties for PIC data
Private mvarAssigningAuthority      As String
Private mvarMaidenName              As String
Private mvarMothersName             As String

'*-- Coleman 101901: Used to get FacilityNCID
Private mobjLogon                   As Logon
Private mDeersID                    As String

Public Property Get DeersID() As String
  DeersID = mDeersID
End Property
Public Property Let DeersID(RHV As String)
  mDeersID = RHV
End Property

Public Property Get ForeignID_Flag() As Boolean

    ForeignID_Flag = mbForeignID

End Property
Public Property Let ForeignID_Flag(newId As Boolean)

    mbForeignID = newId

End Property
'*-- Coleman 051601: Six patient "Flag" properties added
Public Property Get AdoptionFlag() As String
  AdoptionFlag = mvarAdoptionFlag
End Property
Public Property Let AdoptionFlag(RHV As String)
  mvarAdoptionFlag = RHV
End Property

Public Property Get ExpiredFlag() As String
  ExpiredFlag = mvarExpiredFlag
End Property
Public Property Let ExpiredFlag(RHV As String)
  mvarExpiredFlag = RHV
End Property

Public Property Get GuarantorFlag() As String
  GuarantorFlag = mvarGuarantorFlag
End Property
Public Property Let GuarantorFlag(RHV As String)
  mvarGuarantorFlag = RHV
End Property

Public Property Get OrganDonorFlag() As String
  OrganDonorFlag = mvarOrganDonorFlag
End Property
Public Property Let OrganDonorFlag(RHV As String)
  mvarOrganDonorFlag = RHV
End Property

Public Property Get PatientFlag() As String
  PatientFlag = mvarPatientFlag
End Property
Public Property Let PatientFlag(RHV As String)
  mvarPatientFlag = RHV
End Property

Public Property Get UnidentifiedFlag() As String
  UnidentifiedFlag = mvarUnidentifiedFlag
End Property
Public Property Let UnidentifiedFlag(RHV As String)
  mvarUnidentifiedFlag = RHV
End Property

'*-- Coleman 062101: Added two new properties for PIC data
Public Property Get AssigningAuthority() As String
  AssigningAuthority = mvarAssigningAuthority
End Property
Public Property Let AssigningAuthority(RHV As String)
  mvarAssigningAuthority = RHV
End Property

Public Property Get MaidenName() As String
  MaidenName = mvarMaidenName
End Property
Public Property Let MothersName(RHV As String)
  mvarMothersName = RHV
End Property

Public Property Get MothersName() As String
  MothersName = mvarMothersName
End Property
Public Property Let MaidenName(RHV As String)
  mvarMaidenName = RHV
End Property
Public Property Get SponsorSSN() As String
    '*-- Coleman 060801: SCR 13341 - Use GenericData for property
    'SponsorSSN = msSponsorSSN
    SponsorSSN = GetGenericID(NCID_SPONSOR_SSN)
End Property
Public Property Let SponsorSSN(sSponsorSSN As String)
    Dim sSearchString As String
    
    If SponsorSSN <> sSponsorSSN Then
        'msSponsorSSN = sSponsorSSN      '*-- Redundant data, do we want to remove???
        Call UpdateGenericID(NCID_SPONSOR_SSN, sSponsorSSN)
    End If
End Property


'*-- Coleman 061301: Add three new properties
Public Property Get FMPSponsorSSN() As String
    FMPSponsorSSN = GetGenericID(NCID_FMP_SPONSOR_SSN)
End Property
Friend Property Let FMPSponsorSSN(RHV As String)
    If FMPSponsorSSN <> RHV Then
        Call UpdateGenericID(NCID_FMP_SPONSOR_SSN, RHV)
    End If
End Property

'*-- Coleman 061301: Add three new properties
Public Property Get SponsorSrchString() As String
    SponsorSrchString = GetGenericID(NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_SPONSOR_SSN)
End Property
Public Property Let SponsorSrchString(RHV As String)
    If (SponsorSrchString <> RHV) And (Len(RHV) = 5) Then
        Call UpdateGenericID(NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_SPONSOR_SSN, RHV)
    End If
End Property

'*-- Coleman 061301: Add three new properties
Public Property Get PatientSrchString() As String
    PatientSrchString = GetGenericID(NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_PATIENT_SSN)
End Property
Public Property Let PatientSrchString(RHV As String)
    If (PatientSrchString <> RHV) And (Len(RHV) = 5) Then
        Call UpdateGenericID(NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_PATIENT_SSN, RHV)
    End If
End Property


Public Property Get NextGDOrdinal() As Long:     NextGDOrdinal = mvarNextGDOrdinal:   End Property
Public Property Let NextGDOrdinal(RHV As Long):  mvarNextGDOrdinal = RHV:             End Property

Public Property Get UpdateFlag() As String:           UpdateFlag = mvarUpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mvarUpdateFlag = RHV:         End Property


Public Property Get DeployLoc() As String:           DeployLoc = mvarDeployLoc:   End Property
Public Property Let DeployLoc(ByVal RHV As String):  mvarDeployLoc = RHV:         End Property

Public Property Get DeployUnit() As String:          DeployUnit = mvarDeployUnit: End Property
Public Property Let DeployUnit(ByVal RHV As String): mvarDeployUnit = RHV:        End Property

Public Property Get DNBI() As String:                DNBI = mvarDNBI:             End Property
Public Property Let DNBI(ByVal RHV As String):       mvarDNBI = RHV:              End Property


Public Property Get Command_Security() As Collection
  Set Command_Security = mcolCommandSecurity
End Property
Public Property Set Command_Security(RHV As Collection)
  Set mcolCommandSecurity = RHV
End Property
Public Property Let Command_Security(ByVal RHV As Collection)
'////  Call UpdateGenericData(NCID_COMMAND_SECURITY, RHV)
End Property
Public Property Get Command_Securityxx() As Collection
  Set Command_Security = LoadGenericDataCol(NCID_COMMAND_SECURITY)
End Property

Public Property Get Command_Interest() As Collection
  Set Command_Interest = mcolCommandInterest
End Property
Public Property Set Command_Interest(RHV As Collection)
  Set mcolCommandInterest = RHV
End Property
Public Property Let Command_Interest(ByVal RHV As Collection)
 '//// Call UpdateGenericData(NCID_COMMAND_INTEREST, RHV)
End Property
Public Property Get Command_Interestxx() As Collection
    Set Command_Interest = LoadGenericDataCol(NCID_COMMAND_INTEREST)
End Property

'*-- Coleman 100201
Public Property Get Special_Work_Status() As Collection
  Set Special_Work_Status = mcolSpecialWorkStatus
End Property
Public Property Set Special_Work_Status(RHV As Collection)
  Set mcolSpecialWorkStatus = RHV
End Property
Public Property Get Special_Work_Statusxx() As Collection
  Set Special_Work_Status = LoadGenericDataCol(NCID_SPECIAL_WORK_STATUS)
End Property
'New collection class method
Public Property Get Names() As Gems_Names

    Set Names = mobjNames

End Property
Public Property Let Names(PatientName As Gems_Names)


End Property

Public Property Get PatientName() As GEMS_Name
  Set PatientName = mobjName
End Property
Public Property Set PatientName(RHV As GEMS_Name)
  Set mobjName = RHV
End Property

Public Property Get Sponsor() As GEMS_MMI
  Set Sponsor = mobjSponsor
End Property
Public Property Set Sponsor(RHV As GEMS_MMI)
  Set mobjSponsor = RHV
End Property


Public Property Get HomeAddress() As GEMS_Address
  Set HomeAddress = mobjAddresses(NCID_HOME_ADDRESS)
End Property

Public Property Set HomeAddress(ByVal RHV As GEMS_Address)
  RHV.AddrType = NCID_HOME_ADDRESS
  mobjAddresses.AddAddress RHV
End Property

Public Property Get WorkPhone() As GEMS_Phone
  Set WorkPhone = mobjPhones(NCID_WORK_PHONE)
End Property
Public Property Set WorkPhone(ByVal RHV As GEMS_Phone)
  RHV.PhoneType = NCID_WORK_PHONE
  mobjPhones.AddPhone RHV
End Property

Public Property Get ACV_ENROLLMENT_STATUS() As String
  ACV_ENROLLMENT_STATUS = GetGenericData(NCID_ACV_ENROLLMENT_STATUS)
End Property
Public Property Let ACV_ENROLLMENT_STATUS(ByVal RHV As String)
  Call UpdateGenericData(NCID_ACV_ENROLLMENT_STATUS, RHV)
End Property

Public Property Get BirthDate() As Date:                         BirthDate = mvarBirthDate:                           End Property
Public Property Let BirthDate(ByVal RHV As Date):                mvarBirthDate = RHV:                                 End Property

Public Property Get Category() As String
  Category = GetGenericData(NCID_PATCAT)
End Property
Public Property Let Category(ByVal RHV As String)
  Call UpdateGenericData(NCID_PATCAT, RHV)
End Property
Public Property Get Comments() As String
  Comments = GetGenericData(NCID_COMMENTS)
End Property
Public Property Let Comments(ByVal RHV As String)
  Call UpdateGenericData(NCID_COMMENTS, RHV)
End Property

Public Property Get DEERSEligibility() As String
  DEERSEligibility = GetGenericData(NCID_DEERS_ELIGIBILITY)
End Property
Public Property Let DEERSEligibility(ByVal RHV As String)
  Call UpdateGenericData(NCID_DEERS_ELIGIBILITY, RHV)
End Property

Public Property Get DEERSEligibilityEnd() As String
  DEERSEligibilityEnd = GetGenericData(NCID_DEERS_ELIGIBILITY_END_DATE)
End Property
Public Property Let DEERSEligibilityEnd(ByVal RHV As String)
  Call UpdateGenericData(NCID_DEERS_ELIGIBILITY_END_DATE, RHV)
End Property

Public Property Get DEERSEligibilityStart() As String
  DEERSEligibilityStart = GetGenericData(NCID_DEERS_ELIGIBILITY_START_DATE)
End Property
Public Property Let DEERSEligibilityStart(ByVal RHV As String)
  Call UpdateGenericData(NCID_DEERS_ELIGIBILITY_START_DATE, RHV)
End Property

'*-- Coleman 110501: SCR 18301
Public Property Get EMailAddress() As String
  EMailAddress = GetGenericData(NCID_EMAIL)
End Property
Public Property Let EMailAddress(ByVal RHV As String)
  Call UpdateGenericData(NCID_EMAIL, RHV)
End Property


Public Property Get Degree() As String:                          Degree = mobjName.Degree:                                 End Property
Public Property Get Title() As String:                          Title = mobjName.Title:                                  End Property
Public Property Get Suffix() As String:                          Suffix = mobjName.Suffix:                                 End Property

Public Property Get EncounterNumber() As String:                EncounterNumber = mvarEncounterNumber:             End Property
Public Property Let EncounterNumber(ByVal RHV As String):       mvarEncounterNumber = RHV:                          End Property
Public Property Get EncounterFacilityNCID() As String:                EncounterFacilityNCID = mvarEncounterFacilityNCID:             End Property
Public Property Let EncounterFacilityNCID(ByVal RHV As String):       mvarEncounterFacilityNCID = RHV:                          End Property

Public Property Get MedRecordNumber() As String:                MedRecordNumber = mvarMedRecordNumber:             End Property
Public Property Let MedRecordNumber(ByVal RHV As String):       mvarMedRecordNumber = RHV:                          End Property


Public Property Get EnrollmentFacility() As String
  EnrollmentFacility = GetGenericData(NCID_ENROLLMENT_FACILITY)
End Property
Public Property Let EnrollmentFacility(ByVal RHV As String)
  Call UpdateGenericData(NCID_ENROLLMENT_FACILITY, RHV)
End Property

Public Property Get FMP() As String
  '*-- Coleman 090701: SCR 16395
  'FMP = GetGenericData(NCID_PATIENT_FMP)
  FMP = GetGenericID(NCID_PATIENT_FMP)
End Property
Public Property Let FMP(ByVal RHV As String)
  '*-- Coleman 090701: SCR 16395
  'Call UpdateGenericData(NCID_PATIENT_FMP, RHV)
  Call UpdateGenericID(NCID_PATIENT_FMP, RHV)
End Property

Public Property Get FullName() As String
  FullName = mobjName.Last_name & ", " & mobjName.First_name
End Property

Public Property Get MaritalStatus() As String:                   MaritalStatus = mvarMaritalStatus:                   End Property
Public Property Let MaritalStatus(ByVal RHV As String):          mvarMaritalStatus = RHV:                             End Property

Public Property Get Records_Maintained_At() As String
  Records_Maintained_At = GetGenericData(NCID_RECORDS_MAINTAINED_AT)
End Property
Public Property Let Records_Maintained_At(ByVal RHV As String)
  Call UpdateGenericData(NCID_RECORDS_MAINTAINED_AT, RHV)
End Property


Public Property Get NCID() As String
  NCID = mvarNCID
End Property

Public Property Let NCID(ByVal RHV As String)
  mvarNCID = RHV
  mvarUnitNumber = Val(RHV)
End Property

Public Property Get UnitNumber() As Long
  UnitNumber = mvarUnitNumber
End Property

Public Property Let UnitNumber(ByVal RHV As Long)
  mvarNCID = CStr(RHV)
  mvarUnitNumber = RHV
End Property

Public Property Get PCM() As String:                             PCM = mvarPCM:                                       End Property
Public Property Let PCM(ByVal RHV As String):                    mvarPCM = RHV:                                       End Property

Public Property Get PCMAddress() As String:                      PCMAddress = mvarPCMAddress:                         End Property
Public Property Let PCMAddress(ByVal RHV As String):             mvarPCMAddress = RHV:                                End Property

Public Property Get PCMDutyPhone() As GEMS_Phone:                Set PCMDutyPhone = mobjPCMDutyPhone:                 End Property
Public Property Set PCMDutyPhone(ByVal RHV As GEMS_Phone):       Set mobjPCMDutyPhone = RHV:                          End Property

Public Property Get PCMLocation() As String:                     PCMLocation = mvarPCMLocation:                       End Property
Public Property Let PCMLocation(ByVal RHV As String):            mvarPCMLocation = RHV:                               End Property

Public Property Get PCMNCID() As String:                         PCMNCID = mvarPCMNCID:                               End Property
Public Property Let PCMNCID(ByVal RHV As String):                mvarPCMNCID = RHV:                                   End Property

Public Property Get PCMSpecialty() As String:                    PCMSpecialty = mvarPCMSpecialty:                     End Property
Public Property Let PCMSpecialty(ByVal RHV As String):           mvarPCMSpecialty = RHV:                              End Property

Public Property Get Race() As String:                            Race = mvarRace:                                     End Property
Public Property Let Race(ByVal RHV As String):                   mvarRace = RHV:                                      End Property

Public Property Get Rank() As String
  Rank = GetGenericData(NCID_RANK)
End Property
Public Property Let Rank(ByVal RHV As String)
  Call UpdateGenericData(NCID_RANK, RHV)
End Property


Public Property Get Religion() As String:                        Religion = mvarReligion:                             End Property
Public Property Let Religion(ByVal RHV As String):               mvarReligion = RHV:                                  End Property

Public Property Get SEX() As String:                             SEX = mvarSEX:                                       End Property
Public Property Let SEX(ByVal RHV As String):                    mvarSEX = RHV:                                       End Property

Public Property Get SSN() As String
    SSN = mvarSSN
End Property
Public Property Let SSN(ByVal RHV As String)
    '*-- Coleman 071301: Removed redundant code calling UpdateGenericID...
    mvarSSN = RHV
End Property


Public Property Get SWSDiving() As String
  SWSDiving = GetGenericData(NCID_SWS_DIVING_STATUS)
End Property
Public Property Let SWSDiving(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_DIVING_STATUS, RHV)
End Property

Public Property Get SWSFlying() As String
  SWSFlying = GetGenericData(NCID_SWS_FLYING_STATUS)
End Property
Public Property Let SWSFlying(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_FLYING_STATUS, RHV)
End Property

Public Property Get SWSJumping() As String
  SWSJumping = GetGenericData(NCID_SWS_JUMPING_STATUS)
End Property
Public Property Let SWSJumping(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_JUMPING_STATUS, RHV)
End Property

Public Property Get SWSMP() As String
  SWSMP = GetGenericData(NCID_SWS_MP)
End Property
Public Property Let SWSMP(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_MP, RHV)
End Property

Public Property Get SWSOnMobility() As String
  SWSOnMobility = GetGenericData(NCID_SWS_ON_MOBILITY)
End Property
Public Property Let SWSOnMobility(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_ON_MOBILITY, RHV)
End Property


Public Property Get SWSPRP() As String
  SWSPRP = GetGenericData(NCID_SWS_PRP)
End Property
Public Property Let SWSPRP(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_PRP, RHV)
End Property

Public Property Get SWSSubmarine() As String
  SWSSubmarine = GetGenericData(NCID_SWS_SUBMARINE)
End Property
Public Property Let SWSSubmarine(ByVal RHV As String)
  Call UpdateGenericData(NCID_SWS_SUBMARINE, RHV)
End Property


Public Property Get UIC() As String
  UIC = GetGenericData(NCID_UIC)
End Property
Public Property Let UIC(ByVal RHV As String)
  Call UpdateGenericData(NCID_UIC, RHV)
End Property

Public Property Get HomePhone() As GEMS_Phone
  Set HomePhone = mobjPhones(NCID_HOME_PHONE)
End Property
Public Property Set HomePhone(ByVal RHV As GEMS_Phone)
  RHV.PhoneType = NCID_HOME_PHONE
  mobjPhones.AddPhone (RHV)
End Property

Public Property Get MilitaryStatus() As String:                  MilitaryStatus = mvarMilitaryStatus:                 End Property
Public Property Let MilitaryStatus(ByVal RHV As String):         mvarMilitaryStatus = RHV:                            End Property

Public Property Get FacilityNCID() As String:                    FacilityNCID = mvarFacilityNCID:                     End Property
Public Property Let FacilityNCID(ByVal RHV As String):           mvarFacilityNCID = RHV:                              End Property

Public Property Get Addresses() As GEMS_Addresses:          Set Addresses = mobjAddresses: End Property
Public Property Set Addresses(ByVal RHV As GEMS_Addresses): Set mobjAddresses = RHV:       End Property

Public Property Get phones() As GEMS_Phones:                Set phones = mobjPhones:       End Property
Public Property Set phones(ByVal RHV As GEMS_Phones):       Set mobjPhones = RHV:          End Property

Public Property Get Generic_IDs() As GEMS_Generic_IDs:            Set Generic_IDs = mobjGenericIDs:   End Property
Public Property Set Generic_IDs(ByVal RHV As GEMS_Generic_IDs):   Set mobjGenericIDs = RHV:        End Property

Public Property Get Generic_Datas() As GEMS_Generic_Datas:            Set Generic_Datas = mobjGenericDatas:   End Property
Public Property Set Generic_Datas(ByVal RHV As GEMS_Generic_Datas):   Set mobjGenericDatas = RHV:        End Property

'*-- Coleman 052401: Added support of MMI_SITE
Public Property Get Sites() As GEMS_Sites:            Set Sites = mobjSites:   End Property
Public Property Set Sites(ByVal RHV As GEMS_Sites):   Set mobjSites = RHV:     End Property
Private Sub Class_Initialize()

  Set mobjAddresses = New GEMS_Addresses
  Set mobjPhones = New GEMS_Phones
  Set mobjGenericIDs = New GEMS_Generic_IDs
  Set mobjGenericDatas = New GEMS_Generic_Datas
  Set mobjName = New GEMS_Name 'remove soon WR
  Set mobjNames = New Gems_Names 'added WR 5/14/002
  Set mobjSites = New GEMS_Sites            '*-- Coleman 052401
  Set mcolCommandInterest = New Collection
  Set mcolCommandSecurity = New Collection
  '*-- Coleman 100201
  Set mcolSpecialWorkStatus = New Collection
  
  '*-- Coleman 051801: PutPatient assumes that ordinal values are > 0.
  mvarNextGDOrdinal = 1
  
End Sub

Public Sub UpdateGenericData(sNCID As String, sValue As String, Optional sValueNCID As String)
  Dim oGenD As GEMS_Generic_Data

  On Error GoTo ErrorHandler
  Debug.Print "UpdateGenericData(" & sNCID & ", " & sValue & ", " & CStr("" & sValueNCID) & ")"
  
  For Each oGenD In mobjGenericDatas
    If oGenD.Type_NCID = Val(sNCID) Then
      Exit For
    End If
  Next oGenD
  
  If oGenD Is Nothing Then
    Set oGenD = New GEMS_Generic_Data
    oGenD.Type_NCID = sNCID                 '*-- Coleman 060801:
    '*-- Coleman 071301:
    oGenD.Value_Text = sValue
    oGenD.Value_NCID = Val(sValueNCID)
    oGenD.UpdateFlag = "N"
    mobjGenericDatas.AddGenericData oGenD
    'oGenD.Type_NCID = sNCID                '*-- Coleman 060801: Bug???
                                            '*-- Move before ADD???
    'oGenD.Ordinal = mvarNextGDOrdinal          '*-- Coleman 051801
    'mvarNextGDOrdinal = mvarNextGDOrdinal + 1  '*-- Coleman 051801
    If mvarNextGDOrdinal <= oGenD.Ordinal Then  '*-- Coleman 081401
        mvarNextGDOrdinal = oGenD.Ordinal + 1
    Else
        oGenD.Ordinal = mvarNextGDOrdinal
        NextGDOrdinal = NextGDOrdinal + 1
    End If
  Else
    '*-- Coleman 071301:
    oGenD.Value_Text = sValue
    oGenD.Value_NCID = Val(sValueNCID)
    If UpdateFlag = "" Then
        oGenD.UpdateFlag = "U"            'WRogers 8/2001
    Else
        oGenD.UpdateFlag = UpdateFlag
    End If
  End If
  
  Debug.Print "     UpdateGenericData: Type_NCID=" & oGenD.Type_NCID & vbCrLf & _
              "        Value_Text=" & oGenD.Value_Text & vbCrLf & _
              "        Value_NCID=" & oGenD.Value_NCID & vbCrLf & _
              "        Ordinal=" & oGenD.Ordinal & vbCrLf & _
              "        UpdateFlag=" & oGenD.UpdateFlag
                
  Set oGenD = Nothing
  Exit Sub
  
ErrorHandler:
    Exit Sub
End Sub

Public Sub UpdateGenericID(sNCID As String, sValue As String)
  Dim oGenID As GEMS_Generic_Id

  On Error GoTo ErrorHandler
  Debug.Print "UpdateGenericID(" & sNCID & ", " & sValue & ")"
   
  '*-- Look for existing ID object with the same NCID (only finds the
  '*-- first one...there should never be more than one). If one is
  '*-- found, this is an "Update" otherwise it will be a "New".
  For Each oGenID In mobjGenericIDs
    If oGenID.id_type_ncid = sNCID Then
      Exit For
    End If
  Next oGenID
  
  If oGenID Is Nothing Then
    Set oGenID = New GEMS_Generic_Id
    oGenID.id_type_ncid = sNCID             '*-- Coleman 060801:
    '*-- Coleman 071301: SCR 13341
    oGenID.id_value = sValue
    oGenID.UpdateFlag = "N"
    mobjGenericIDs.AddGenericID oGenID
    'oGenID.id_type_ncid = sNCID            '*-- Coleman 060801: Bug???
                                            '*-- Move before ADD???
  Else
    '*-- Coleman 060801: SCR 13341
    '*-- Coleman 071301: SCR 13341 fix removed, problem resolved elsewhere.
    oGenID.id_value = sValue
    If UpdateFlag = "" Then
        oGenID.UpdateFlag = "U"                 '--- Wang SCR16676 10/19/01
    Else
        oGenID.UpdateFlag = UpdateFlag
    End If
  End If

  Set oGenID = Nothing
  Exit Sub
  
ErrorHandler:
    Exit Sub
End Sub


Public Function GetGenericData(Optional NCID As String) As String
  Dim oGenD As GEMS_Generic_Data
  On Error Resume Next
  Set oGenD = mobjGenericDatas(NCID)
  If oGenD Is Nothing Then
    GetGenericData = ""
    Exit Function
  Else
    GetGenericData = oGenD.Value_Text
  End If
  Set oGenD = Nothing
  
End Function

'*-- Coleman 090401: Support for Property request by NCID
Public Function GetGenericDataValueNCID(Optional NCID As String) As String
  Dim oGenD As GEMS_Generic_Data
  On Error Resume Next
  Set oGenD = mobjGenericDatas(NCID)
  If oGenD Is Nothing Then
    GetGenericDataValueNCID = ""
    Exit Function
  Else
    GetGenericDataValueNCID = oGenD.Value_NCID
  End If
  Set oGenD = Nothing
  
End Function

Public Function GetGenericID(Optional NCID As String) As String
  Dim oGenID As GEMS_Generic_Id
  On Error Resume Next
  Set oGenID = mobjGenericIDs(NCID)
  If oGenID Is Nothing Then
    GetGenericID = ""
    Exit Function
  Else
    GetGenericID = oGenID.id_value
  End If
  Set oGenID = Nothing
  
End Function

Public Function SaveSiteObj(Optional FacilityNCID As String, _
                            Optional MedRecNum As String = "00000", _
                            Optional RadiologyNum As String = "000") As Boolean
    Dim oSite As GEMS_Site
    
    Set oSite = New GEMS_Site
    
    '*-- Coleman 103001: Normally FacilityNCID should be passed in since
    '*-- it has to be read out of the OSH_INI_Settings table.
    '*-- Default: Force FacilityNCID to local Theater facility from Logon.
    If FacilityNCID = "" Then
        If mobjLogon Is Nothing Then
            Set mobjLogon = New Logon
        End If
        FacilityNCID = mobjLogon.FacilityNCID
    End If
    oSite.FacilityNCID = FacilityNCID
    oSite.MedRecNum = MedRecNum
    oSite.RadiologyNum = RadiologyNum
    oSite.UpdateFlag = UpdateFlag
    Sites.AddSite oSite
    MedRecordNumber = oSite.MedRecNum

End Function

Private Sub Class_Terminate()

  Set mobjAddresses = Nothing
  Set mobjPhones = Nothing
  Set mobjGenericIDs = Nothing
  Set mobjGenericDatas = Nothing
  Set mobjSites = Nothing           '*-- Coleman 052401
  Set mobjSponsor = Nothing
  Set mcolCommandInterest = Nothing
  Set mcolCommandSecurity = Nothing
  '*-- Coleman 100201
  Set mcolSpecialWorkStatus = Nothing

End Sub

'*-- Coleman 100201: Copied from Security_Interestxx() and made generic.
'*-- To Do: simplify control logic...
Private Function LoadGenericDataCol(sNCID As String) As Collection
    Dim str As String
    Dim mCol As Collection
    Set mCol = New Collection
    
    str = GetGenericData(sNCID)
    If Not str = "" Then
      mCol.Add str
    Else
      Set LoadGenericDataCol = mCol
      Exit Function
    End If
    Do While Not str = ""
        str = GetGenericData()
        If Not str = "" Then
            mCol.Add str
        End If
    Loop
    Set LoadGenericDataCol = mCol
End Function

