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
Attribute VB_Name = "EncounterProcedureRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sDiagnosisDataId As String
Private sProcedureNCID As String
Private sTypeNCID As String
Private sDataId As String
Private sEncounterDataId As String
Private sName As String
Private sCode As String

Public Property Get DiagnosisDataId() As String
    DiagnosisDataId = sDiagnosisDataId
End Property

Public Property Let DiagnosisDataId(ByVal sNewValue As String)
    sDiagnosisDataId = sNewValue
End Property

Public Property Get ProcedureNCID() As String
    ProcedureNCID = sProcedureNCID
End Property

Public Property Let ProcedureNCID(ByVal sNewValue As String)
     sProcedureNCID = sNewValue
End Property

Public Property Get TypeNCID() As String
    TypeNCID = sTypeNCID
End Property

Public Property Let TypeNCID(ByVal sNewValue As String)
    sTypeNCID = sNewValue
End Property

Public Property Get DataId() As String
    DataId = sDataId
End Property

Public Property Let DataId(ByVal sNewValue As String)
    sDataId = sNewValue
End Property

Public Property Get EncounterDataId() As String
    EncounterDataId = sEncounterDataId
End Property

Public Property Let EncounterDataId(ByVal sNewValue As String)
    sEncounterDataId = sNewValue
End Property

Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Name(ByVal sNewValue As String)
    sName = sNewValue
End Property

Public Property Get Code() As String
    Code = sCode
End Property

Public Property Let Code(ByVal sNewValue As String)
    sCode = sNewValue
End Property
