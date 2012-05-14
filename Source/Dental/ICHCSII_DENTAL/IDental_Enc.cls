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
Attribute VB_Name = "IDental_Enc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get DentalEncounterId() As Long

End Property
Public Property Get MedicalEncounterId() As Long

End Property

Public Property Get PatientId() As Long

End Property

Public Property Get EncounterDate() As Date

End Property

Public Property Get EligibilityFlag() As Boolean

End Property
Public Property Get EligibilityDate() As Date

End Property
Public Property Get SBE() As IntrfcCHCSII_Dental.IDENTAL_SBE

End Property

Public Property Get DRC() As IDENTAL_DRC

End Property
Public Function LoadDentalEncounter(ByVal lngPatientID As Long, ByVal lngEncounterId As Long, _
            ByVal lngFacilityNCID As Long, ByVal dteApptDate As Date) As Boolean

End Function

Public Property Get NewEncounter() As Boolean

End Property

