VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MedWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWarningType As String
Private mWarningText As String
Private mDrugClass As String
Private mDrugList As String
Private mSeverityLevel As String

Private mMonographID As String
Private mMonographType As Integer  'fdbMonographType



Public Property Get SeverityLevel() As String
  SeverityLevel = mSeverityLevel
End Property
Public Property Let SeverityLevel(sValue As String)
  mSeverityLevel = sValue
End Property

Public Property Get MonographType() As Integer ' fdbMonographType
  MonographType = mMonographType
End Property
Public Property Let MonographType(sValue As Integer) ' fdbMonographType)
  mMonographType = sValue
End Property

Public Property Get WarningType() As String
  WarningType = mWarningType
End Property
Public Property Let WarningType(sValue As String)
  mWarningType = sValue
End Property

Public Property Get WarningText() As String
  WarningText = mWarningText
End Property
Public Property Let WarningText(sValue As String)
  mWarningText = sValue
End Property

Public Property Let DrugClass(sValue As String)
  mDrugClass = sValue
End Property
Public Property Get DrugClass() As String
  DrugClass = mDrugClass
End Property

Public Property Let DrugList(sValue As String)
  mDrugList = sValue
End Property
Public Property Get DrugList() As String
  DrugList = mDrugList
End Property

Public Property Let MonographID(sValue As String)
  mMonographID = sValue
End Property
Public Property Get MonographID() As String
  MonographID = mMonographID
End Property

