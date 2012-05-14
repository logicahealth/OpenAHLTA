VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInitUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Property Set CWComm(oComm As Comm)

    Set cwobjCWComm = oComm
    
End Property


Public Property Let MyProgID(ProgID As Long)
    
    cwlMyProgId = ProgID
    
End Property

Public Property Let PatientAge(ByVal intAge As Integer)
    gintPatientAge = intAge
End Property

Public Property Get PatientAge() As Integer
    PatientAge = gintPatientAge
End Property
