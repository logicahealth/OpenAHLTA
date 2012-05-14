VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "thr_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moGemsMMI       As GEMS_MMI
Private moAddress       As CHCSII_MMI.thr_Address
Private moPhone         As CHCSII_MMI.thr_Phone
Private moFacility      As CHCSII_MMI.thr_Facility
Private moGenericData   As CHCSII_MMI.thr_GenericData
Private moGenericID     As CHCSII_MMI.thr_GenericId
Private moName          As CHCSII_MMI.thr_Name
Public bSponsor         As Boolean

Implements IMember
Implements IMilMember   'added to support patient attributes not in Member WR 5/1/2002

'Private mlUnitNumber  As Long

Friend Property Set GemsMMI(gMMI As GEMS_MMI)

    Set moGemsMMI = gMMI
    
    With gMMI
        Set moAddress = New thr_Address
        Set moAddress.GemsAddress = .Addresses
        Set moPhone = New thr_Phone
        Set moPhone.GemsPhone = .phones
        Set moFacility = New thr_Facility
        Set moFacility.GemsFacility = .Sites
        Set moGenericData = New thr_GenericData
        Set moGenericData.GemsGenericData = .Generic_Datas
        Set moGenericID = New thr_GenericId
        Set moGenericID.SetGenericID = .Generic_IDs
        Set moName = New thr_Name
        Set moName.gemsName = .Names
    End With

End Property

Private Function IMember_Add() As Long
    'Debug.Assert 0
End Function


Private Property Let IMember_Address(ByVal RHS As Variant)

    Dim iAddr       As CHCSII_MMILib.IAddress
    Dim gemsAddr    As GEMS_Address

    If moAddress Is Nothing Then
        Set gemsAddr = New GEMS_Address
        gemsAddr.UpdateFlag = "N"
        gemsAddr.AddrType = RHS
        moGemsMMI.Addresses.AddAddress gemsAddr
        Set moAddress = New CHCSII_MMI.thr_Address
        Set moAddress.GemsAddress = moGemsMMI.Addresses
    End If
    
    Set iAddr = moAddress

    iAddr.Type = RHS

End Property

Private Property Get IMember_Address() As Variant
    Dim iAddr As CHCSII_MMILib.IAddress

    If moAddress Is Nothing And moGemsMMI.Addresses.Count > 0 Then
        Set moAddress = New CHCSII_MMI.thr_Address
        Set moAddress.GemsAddress = moGemsMMI.Addresses
    ElseIf moAddress Is Nothing Then
        Set IMember_Address = Nothing
        Exit Property
    End If
    
    Set iAddr = moAddress
    Set IMember_Address = iAddr

End Property

Private Property Let IMember_Adopted(ByVal RHS As String)

    moGemsMMI.AdoptionFlag = RHS

End Property

Private Property Get IMember_Adopted() As String

    IMember_Adopted = moGemsMMI.AdoptionFlag

End Property

Private Property Let IMember_AppVersion(ByVal RHS As String)

    'Debug.Assert 0
End Property

Private Property Get IMember_AppVersion() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_BirthDate(ByVal RHS As Date)

    moGemsMMI.BirthDate = RHS

End Property

Private Property Get IMember_BirthDate() As Date

    IMember_BirthDate = moGemsMMI.BirthDate

End Property


Private Function IMember_Bobbit() As Boolean
    'Debug.Assert 0
End Function

Private Property Get IMember_BOF() As Boolean
    'Debug.Assert 0
End Property

Private Property Let IMember_Confidence(ByVal RHS As Long)
    'Debug.Assert 0
End Property

Private Property Get IMember_Confidence() As Long
    'Debug.Assert 0
End Property

Private Property Let IMember_CreateApp(ByVal RHS As String)

End Property

Private Property Get IMember_CreateApp() As String

End Property

Private Property Let IMember_Creator(ByVal RHS As Long)
    'Debug.Assert 0
End Property

Private Property Get IMember_Creator() As Long
    'Debug.Assert 0
End Property

Private Property Let IMember_CtreateTime(ByVal RHS As Date)
    'Debug.Assert 0
End Property

Private Property Get IMember_CtreateTime() As Date
    'Debug.Assert 0
End Property

Private Property Let IMember_DateOfDeath(ByVal RHS As Date)

    'moGemsMMI.ExpiredFlag
    'Debug.Assert 0
End Property

Private Property Get IMember_DateOfDeath() As Date
    'Debug.Assert 0
End Property

Private Property Let IMember_Deactivate(ByVal RHS As String)
    'Debug.Assert 0
End Property

Private Property Get IMember_Deactivate() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_DOBEstimated(ByVal RHS As String)
    'Debug.Assert 0
End Property

Private Property Get IMember_DOBEstimated() As String
    'Debug.Assert 0
End Property

Private Property Get IMember_EOF() As Boolean

    'Debug.Assert 0
End Property

Private Property Let IMember_Expired(ByVal RHS As String)
    
    moGemsMMI.ExpiredFlag = RHS

End Property

Private Property Get IMember_Expired() As String

    IMember_Expired = moGemsMMI.ExpiredFlag

End Property

Private Property Let IMember_Facility(ByVal RHS As Variant)

    Dim iFac            As CHCSII_MMILib.IFacility
    Dim GemsFacility    As GEMS_Site
    
    If moFacility Is Nothing Then
        Dim gemsSite As GEMS_Site
        Set gemsSite = New GEMS_Site
        gemsSite.UpdateFlag = "N"
        moGemsMMI.Sites.AddSite gemsSite
        Set moFacility = New CHCSII_MMI.thr_Facility
        Set moFacility.GemsFacility = moGemsMMI.Sites
    End If
    
    Set iFac = moFacility

    iFac.Type = RHS

End Property

Private Property Get IMember_Facility() As Variant
    Dim iFac As CHCSII_MMILib.IFacility

    If moFacility Is Nothing And moGemsMMI.Sites.Count > 0 Then
        Set moFacility = New CHCSII_MMI.thr_Facility
        Set moFacility.GemsFacility = moGemsMMI.Sites
            
    ElseIf moFacility Is Nothing Then
        Set IMember_Facility = Nothing
        Exit Property
    End If
    
    Set iFac = moFacility
    Set IMember_Facility = iFac
    
End Property

Private Property Let IMember_GenericData(ByVal RHS As Variant)

    Dim oGenData As CHCSII_MMILib.IGenericData
    Dim GemsGenericData    As GEMS_Generic_Data
    
    If moGenericData Is Nothing Then
        Dim gemData As GEMS_Generic_Data
        Set gemData = New GEMS_Generic_Data
        gemData.UpdateFlag = "N"
        moGemsMMI.Generic_Datas.AddGenericData gemData
        Set moGenericData = New CHCSII_MMI.thr_GenericData
        Set moGenericData.GemsGenericData = moGemsMMI.Generic_Datas
        
    End If

    Set oGenData = moGenericData
    oGenData.Type = RHS

End Property

Private Property Get IMember_GenericData() As Variant
    Dim iGenMMI As CHCSII_MMILib.IGenericData
    
    If moGenericData Is Nothing And moGemsMMI.Generic_Datas.Count > 0 Then
        Set moGenericData = New CHCSII_MMI.thr_GenericData
        
        Set moGenericData.GemsGenericData = moGemsMMI.Generic_Datas
    ElseIf moGenericData Is Nothing Then
        Set IMember_GenericData = Nothing
        Exit Property
    End If

    Set iGenMMI = moGenericData
    Set IMember_GenericData = iGenMMI

End Property

Private Property Let IMember_GenericId(ByVal RHS As Variant)

    Dim oGenData        As CHCSII_MMILib.IGenericId
    Dim gemsGenericId   As GEMS_Generic_Id
    
    If moGenericID Is Nothing Then
        Dim gemsID As GEMS_Generic_Id
        
        Set gemsID = New GEMS_Generic_Id
        gemsID.UpdateFlag = "N"
        moGemsMMI.Generic_IDs.AddGenericID gemsID
        Set moGenericID = New CHCSII_MMI.thr_GenericId
        Set moGenericID.SetGenericID = moGemsMMI.Generic_IDs
    End If
    
    Set oGenData = moGenericID
    Debug.Assert Not oGenData Is Nothing

    oGenData.Type = RHS
    
End Property

Private Property Get IMember_GenericId() As Variant
    Dim oGenId As CHCSII_MMILib.IGenericId
    
    If moGenericID Is Nothing And moGemsMMI.Generic_IDs.Count > 0 Then
        Dim gemsID As GEMS_Generic_Id
        Set gemsID = New GEMS_Generic_Id
        
        Set moGenericID = New CHCSII_MMI.thr_GenericId
        
        gemsID.UpdateFlag = "N"
        moGemsMMI.Generic_IDs.AddGenericID gemsID
        
        Set moGenericID.SetGenericID = moGemsMMI.Generic_IDs
        
    ElseIf moGenericID Is Nothing Then
        Set IMember_GenericId = Nothing
        Exit Property
    End If
    
    Set oGenId = moGenericID
    
    Set IMember_GenericId = oGenId

End Property
Friend Property Get GemsMMI() As GEMS_MMI

    Set GemsMMI = moGemsMMI

End Property
Private Sub IMember_GotoFirst()
    'Debug.Assert 0
End Sub

Private Sub IMember_GotoLast()
    'Debug.Assert 0
End Sub

Private Sub IMember_GotoNext()
    'Debug.Assert 0
End Sub

Private Sub IMember_GotoPrevious()
    'Debug.Assert 0
End Sub

Private Property Let IMember_GovtHealthNumber(ByVal RHS As String)
    'Debug.Assert 0
End Property

Private Property Get IMember_GovtHealthNumber() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_Guarantor(ByVal RHS As String)
    
    moGemsMMI.GuarantorFlag = RHS
    
End Property

Private Property Get IMember_Guarantor() As String

    IMember_Guarantor = moGemsMMI.GuarantorFlag
    
End Property

Private Property Let IMember_Institution(ByVal RHS As String)
    'Debug.Assert 0
End Property
'Could this be our theater "facility" or EncounterFacility?
'should be a Concept Control Representation of Institution
Private Property Get IMember_Institution() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_MaidenName(ByVal RHS As String)

    moGemsMMI.MaidenName = RHS

End Property

Private Property Get IMember_MaidenName() As String

    IMember_MaidenName = moGemsMMI.MaidenName

End Property

Private Property Let IMember_MaritalStatus(ByVal RHS As String)

    moGemsMMI.MaritalStatus = RHS
    
End Property

Private Property Get IMember_MaritalStatus() As String

    IMember_MaritalStatus = moGemsMMI.MaritalStatus
    
End Property

Private Property Let IMember_MMIVersion(ByVal RHS As String)
    'Debug.Assert 0
End Property

Private Property Get IMember_MMIVersion() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_MomMaidenName(ByVal RHS As String)

    moGemsMMI.MothersName = RHS

End Property

Private Property Get IMember_MomMaidenName() As String

    IMember_MomMaidenName = moGemsMMI.MothersName

End Property

Private Property Let IMember_Name(ByVal RHS As Variant)

    Dim iNM         As CHCSII_MMILib.IName
    Dim gemsName    As GEMS_Name

    If moName Is Nothing Then
        'Dim gemsName As GEMS_Name
        Set gemsName = New GEMS_Name
        gemsName.UpdateFlag = "N"
        moGemsMMI.Names.AddName gemsName
        Set moName = New CHCSII_MMI.thr_Name
        Set moName.gemsName = moGemsMMI.Names
    End If
    
    Set iNM = moName

    iNM.Type = RHS
    
End Property

Private Property Get IMember_Name() As Variant
    Dim iNM As CHCSII_MMILib.IName
    
    If moName Is Nothing And moGemsMMI.PatientName Is Nothing Then
        Set moName = New CHCSII_MMI.thr_Name
        Set moName.gemsName = moGemsMMI.PatientName
    ElseIf moName Is Nothing Then
        Set IMember_Name = Nothing
        Exit Property
    End If

    Set iNM = moName
    Set IMember_Name = iNM

End Property

Private Property Let IMember_OrganDonor(ByVal RHS As String)

    moGemsMMI.OrganDonorFlag = RHS

End Property

Private Property Get IMember_OrganDonor() As String

    IMember_OrganDonor = moGemsMMI.OrganDonorFlag

End Property

Private Property Let IMember_Passport(ByVal RHS As String)

        'Debug.Assert 0
End Property

Private Property Get IMember_Passport() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_Patient(ByVal RHS As String)
    ' SCR #45685;   Developer:  01/06/2004
    moGemsMMI.PatientFlag = RHS
    'Debug.Assert 0
End Property

Private Property Get IMember_Patient() As String
     ' SCR #45685;   Developer:  01/06/2004
     IMember_Patient = moGemsMMI.PatientFlag
    'Debug.Assert 0
End Property

Private Property Let IMember_Phone(ByVal RHS As Variant)

    Dim iPH As CHCSII_MMILib.IPhone
    Dim gemsPh As GEMS_Phone
    
    
    If moPhone Is Nothing Then
        'Dim gemsPh As GEMS_Phone
        Set gemsPh = New GEMS_Phone
        gemsPh.UpdateFlag = "N"
        moGemsMMI.phones.AddPhone gemsPh
        Set moPhone = New CHCSII_MMI.thr_Phone
        Set moPhone.GemsPhone = moGemsMMI.phones
    End If
    
    Set iPH = moPhone

    iPH.Type = RHS
    
End Property

Private Property Get IMember_Phone() As Variant
    Dim iPH  As CHCSII_MMILib.IPhone

    If moPhone Is Nothing And moGemsMMI.phones.Count > 0 Then
        Set moPhone = New CHCSII_MMI.thr_Phone
        Set moPhone.GemsPhone = moGemsMMI.phones
    ElseIf moPhone Is Nothing Then
        Set IMember_Phone = Nothing
        Exit Property
    End If

    Set iPH = moPhone
    Set IMember_Phone = iPH

End Property

Private Property Let IMember_Race(ByVal RHS As String)

    moGemsMMI.Race = RHS

End Property

Private Property Get IMember_Race() As String

    IMember_Race = moGemsMMI.Race

End Property

Private Property Let IMember_ReferenceId(ByVal RHS As String)
    moGemsMMI.DeersID = RHS
End Property

Private Property Get IMember_ReferenceId() As String
    IMember_ReferenceId = moGemsMMI.DeersID
End Property


Private Function IMember_ReleaseLock() As Boolean
    'Debug.Assert 0
End Function

Private Property Let IMember_Religion(ByVal RHS As String)

    moGemsMMI.Religion = RHS

End Property

Private Property Get IMember_Religion() As String

    IMember_Religion = moGemsMMI.Religion

End Property


Private Function IMember_Remove() As Boolean
    'Debug.Assert 0
End Function

Private Property Let IMember_Self(ByVal RHS As Variant)
    'Debug.Assert 0
End Property

Private Property Get IMember_Self() As Variant
    'Debug.Assert 0
End Property

Private Property Let IMember_Sex(ByVal RHS As String)

    moGemsMMI.Sex = RHS

End Property

Private Property Get IMember_Sex() As String

    IMember_Sex = moGemsMMI.Sex

End Property

Private Property Let IMember_SSN(ByVal RHS As String)

    moGemsMMI.SSN = RHS

End Property

Private Property Get IMember_SSN() As String

    IMember_SSN = moGemsMMI.SSN

End Property

Private Property Let IMember_SSNCorrect(ByVal RHS As String)
    'Debug.Assert 0
End Property

Private Property Get IMember_SSNCorrect() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_SSNMissingReason(ByVal RHS As String)
    'Debug.Assert 0
End Property

Private Property Get IMember_SSNMissingReason() As String
    'Debug.Assert 0
End Property

Private Property Let IMember_Unidentified(ByVal RHS As String)

    moGemsMMI.UnidentifiedFlag = RHS

End Property

Private Property Get IMember_Unidentified() As String

    IMember_Unidentified = moGemsMMI.UnidentifiedFlag

End Property

Private Property Get IMember_UnitNumber() As Long

    IMember_UnitNumber = moGemsMMI.UnitNumber

End Property

Private Property Let IMember_UnitNumber(ByVal RHS As Long)

    moGemsMMI.UnitNumber = RHS

End Property

Private Function IMember_WriteUpdate() As Boolean
    Dim oGemsOps As GEMS_MMIOps
    
    'save our data
    Set oGemsOps = New GEMS_MMIOps

    IMember_WriteUpdate = oGemsOps.PutPatient(moGemsMMI)

End Function

Private Property Let IMilMember_DeployLocation(ByVal RHS As String)

    moGemsMMI.DeployLoc = RHS

End Property

Private Property Get IMilMember_DeployLocation() As String

    IMilMember_DeployLocation = moGemsMMI.DeployLoc

End Property

Private Property Let IMilMember_DeployUnit(ByVal RHS As String)

    moGemsMMI.DeployUnit = RHS

End Property

Private Property Get IMilMember_DeployUnit() As String

    IMilMember_DeployUnit = moGemsMMI.DeployUnit

End Property

Private Property Let IMilMember_DNBI(ByVal RHS As String)

    moGemsMMI.DNBI = RHS

End Property

Private Property Get IMilMember_DNBI() As String

    IMilMember_DNBI = moGemsMMI.DNBI

End Property
