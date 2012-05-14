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
Attribute VB_Name = "thr_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IUser

Private Property Let IUser_ActivateState(ByVal RHS As String)

End Property

Private Property Get IUser_ActivateState() As String

End Property

Private Function IUser_Add() As Long

End Function

Private Property Let IUser_Address(ByVal RHS As Variant)

End Property

Private Property Get IUser_Address() As Variant

End Property

Private Property Let IUser_Adopted(ByVal RHS As String)

End Property

Private Property Get IUser_Adopted() As String

End Property

Private Property Let IUser_AppVersion(ByVal RHS As String)

End Property

Private Property Get IUser_AppVersion() As String

End Property

Private Property Let IUser_BirthDate(ByVal RHS As Date)

End Property

Private Property Get IUser_BirthDate() As Date

End Property

Private Function IUser_Bobbit() As Boolean

End Function


Private Property Get IUser_BOF() As Boolean

End Property

Private Property Let IUser_Confidence(ByVal RHS As Long)

End Property

Private Property Get IUser_Confidence() As Long

End Property

Private Property Let IUser_CreateApp(ByVal RHS As String)

End Property

Private Property Get IUser_CreateApp() As String

End Property

Private Property Let IUser_Creator(ByVal RHS As Long)

End Property

Private Property Get IUser_Creator() As Long

End Property

Private Property Let IUser_CtreateTime(ByVal RHS As Date)

End Property

Private Property Get IUser_CtreateTime() As Date

End Property

Private Property Let IUser_DateOfDeath(ByVal RHS As Date)

End Property

Private Property Get IUser_DateOfDeath() As Date

End Property

Private Property Let IUser_Deactivate(ByVal RHS As String)

End Property

Private Property Get IUser_Deactivate() As String

End Property

Private Property Let IUser_DEANumber(ByVal RHS As String)

End Property

Private Property Get IUser_DEANumber() As String

End Property

Private Property Let IUser_DOBEstimated(ByVal RHS As String)

End Property

Private Property Get IUser_DOBEstimated() As String

End Property


Private Property Get IUser_EOF() As Boolean

End Property

Private Property Let IUser_Expired(ByVal RHS As String)

End Property

Private Property Get IUser_Expired() As String

End Property

Private Property Let IUser_Facility(ByVal RHS As Variant)

End Property

Private Property Get IUser_Facility() As Variant

End Property

Private Sub IUser_GotoFirst()

End Sub

Private Sub IUser_GotoLast()

End Sub

Private Sub IUser_GotoNext()

End Sub

Private Sub IUser_GotoPrevious()

End Sub

Private Property Let IUser_GovtHealthNumber(ByVal RHS As String)

End Property

Private Property Get IUser_GovtHealthNumber() As String

End Property

Private Property Let IUser_Guarantor(ByVal RHS As String)

End Property

Private Property Get IUser_Guarantor() As String

End Property

Private Property Let IUser_InactivateDate(ByVal RHS As Date)

End Property

Private Property Get IUser_InactivateDate() As Date

End Property

Private Property Let IUser_Initials(ByVal RHS As String)

End Property

Private Property Get IUser_Initials() As String

End Property

Private Property Let IUser_Institution(ByVal RHS As String)

End Property

Private Property Get IUser_Institution() As String

End Property

Private Property Let IUser_LastLogon(ByVal RHS As Date)

End Property

Private Property Get IUser_LastLogon() As Date

End Property

Private Property Let IUser_Location(ByVal RHS As String)

End Property

Private Property Get IUser_Location() As String

End Property

Private Property Let IUser_LogonName(ByVal RHS As String)

End Property

Private Property Get IUser_LogonName() As String

End Property

Private Property Let IUser_MaidenName(ByVal RHS As String)

End Property

Private Property Get IUser_MaidenName() As String

End Property

Private Property Let IUser_MaritalStatus(ByVal RHS As String)

End Property

Private Property Get IUser_MaritalStatus() As String

End Property

Private Property Let IUser_MMIVersion(ByVal RHS As String)

End Property

Private Property Get IUser_MMIVersion() As String

End Property

Private Property Let IUser_MomMaidenName(ByVal RHS As String)

End Property

Private Property Get IUser_MomMaidenName() As String

End Property

Private Property Let IUser_Name(ByVal RHS As Variant)

End Property

Private Property Get IUser_Name() As Variant

End Property

Private Property Let IUser_OrganDonor(ByVal RHS As String)

End Property

Private Property Get IUser_OrganDonor() As String

End Property

Private Property Let IUser_Passport(ByVal RHS As String)

End Property

Private Property Get IUser_Passport() As String

End Property

Private Property Let IUser_Patient(ByVal RHS As String)

End Property

Private Property Get IUser_Patient() As String

End Property

Private Property Let IUser_Phone(ByVal RHS As Variant)

End Property

Private Property Get IUser_Phone() As Variant

End Property

Private Property Let IUser_Preferences(ByVal RHS As String)

End Property

Private Property Get IUser_Preferences() As String

End Property

Private Property Let IUser_Race(ByVal RHS As String)

End Property

Private Property Get IUser_Race() As String

End Property

Private Property Let IUser_ReferenceId(ByVal RHS As String)

End Property

Private Property Get IUser_ReferenceId() As String

End Property

Private Property Let IUser_Religion(ByVal RHS As String)

End Property

Private Property Get IUser_Religion() As String

End Property

Private Function IUser_Remove() As Boolean

End Function

Private Property Let IUser_Role(ByVal RHS As String)

End Property

Private Property Get IUser_Role() As String

End Property

Private Property Let IUser_Self(ByVal RHS As Variant)

End Property

Private Property Get IUser_Self() As Variant

End Property

Private Property Let IUser_Sex(ByVal RHS As String)

End Property

Private Property Get IUser_Sex() As String

End Property

Private Property Let IUser_Signature(ByVal RHS As String)

End Property

Private Property Get IUser_Signature() As String

End Property

Private Property Let IUser_SiteId(ByVal RHS As String)

End Property

Private Property Get IUser_SiteId() As String

End Property

Private Property Let IUser_Specialty(ByVal RHS As String)

End Property

Private Property Get IUser_Specialty() As String

End Property

Private Property Let IUser_SSN(ByVal RHS As String)

End Property

Private Property Get IUser_SSN() As String

End Property

Private Property Let IUser_SSNCorrect(ByVal RHS As String)

End Property

Private Property Get IUser_SSNCorrect() As String

End Property

Private Property Let IUser_SSNMissingReason(ByVal RHS As String)

End Property

Private Property Get IUser_SSNMissingReason() As String

End Property

Private Property Let IUser_StateLicense(ByVal RHS As String)

End Property

Private Property Get IUser_StateLicense() As String

End Property

Private Property Let IUser_Unidentified(ByVal RHS As String)

End Property

Private Property Get IUser_Unidentified() As String

End Property

Private Property Let IUser_UnitNumber(ByVal RHS As Long)

End Property

Private Property Get IUser_UnitNumber() As Long

End Property

Private Property Let IUser_UserNcid(ByVal RHS As String)

End Property

Private Property Get IUser_UserNcid() As String

End Property

Private Property Let IUser_UserNumber(ByVal RHS As String)

End Property

Private Property Get IUser_UserNumber() As String

End Property

Private Function IUser_WriteUpdate() As Boolean

End Function
