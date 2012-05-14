VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrintInpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public RegNum               As String
Public Specialty            As String
Public Capacity             As String
Public Bed                  As String
Public Rank                 As String
Public Patient              As String
Public FMPSSN               As String
Public Unit                 As String
Public PatService           As String
Public AdmittingDiagnosis   As String
Public AdmissionDate        As String
Public AdmissionComments    As String
Public Day                  As String
Public Location             As String
Public Status               As String
Public DischargeDate        As String
Public Disposition          As String
Public DischargeDiagnosis   As String
Public AdmittingProvider    As String
Public AttendingProvider    As String
Public TransferDate         As String
Public Service              As String

