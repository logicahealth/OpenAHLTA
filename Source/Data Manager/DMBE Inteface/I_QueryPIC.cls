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
Attribute VB_Name = "I_QueryPIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PicStatus() As DMStatus: End Property
Public Property Let PicStatus(Sts As DMStatus): End Property

Public Property Get PatientName() As String: End Property
Public Property Let PatientName(name As String): End Property

Public Property Get PatientSSN() As String: End Property
Public Property Let PatientSSN(SSN As String): End Property

Public Property Get PatientUnitNumber() As String: End Property
Public Property Let PatientUnitNumber(UnitNumber As String): End Property

Public Property Get FileName() As String: End Property
Public Property Let FileName(sValue As String): End Property

Public Property Get PICDataSetExists(Identifier As Variant) As Boolean: End Property
Public Property Let PICDataSetExists(Identifier As Variant, DSExists As Boolean): End Property

Public Sub InitVariables(): End Sub

Public Property Get DataSetsInPIC() As Variant

End Property

Public Property Let DataSetsInPIC(ByVal DataSetLongArray As Variant)

End Property

Public Property Get AGGPICPatients() As Variant

End Property

Public Property Let AGGPICPatients(ByVal enmDMFPatientInfo As Variant)

End Property

Public Property Get PicFormat() As DMFormats: End Property
Public Property Let PicFormat(Sts As DMFormats): End Property

