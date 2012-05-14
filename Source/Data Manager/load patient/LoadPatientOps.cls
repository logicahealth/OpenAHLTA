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
Attribute VB_Name = "LoadPatientOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function LoadNewPatientProperties( _
    sUnit_Number As String, _
    Optional sSponsorSSN As String, Optional sFMP As String, _
    Optional sSurrogateKey, Optional sLastName As String, _
    Optional sFirstName As String, Optional sMiddleName As String, _
    Optional sPatientCadence As String, Optional sSSN As String, _
    Optional sDOB As String, Optional sGender As String, _
    Optional sRace As String, Optional sEthnicity As String, _
    Optional sMaritalStatus As String, Optional sPatientCategory As String, _
    Optional sOfficialRecordInd As String, Optional sServiceBranch As String, _
    Optional sRank As String, Optional sPayGrade As String, _
    Optional sHomeUnitUIC As String, Optional sMos As String, _
    Optional sFlightStatus As String, _
    Optional sPresidentialProgramDuty As String, _
    Optional sPrpQual As String, Optional sSCISecurityClearance As String, _
    Optional sAssigningAuthority As String, _
    Optional sMothersMaidenName As String) As Boolean

End Function

