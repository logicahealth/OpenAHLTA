VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MMI_Site"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarFacilityNCID As String
Private mvarMedRecNum As String
Private mvarRadiologyNum As String

Public Property Get FacilityNCID() As String
    FacilityNCID = mvarFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal RHV As String)
    mvarFacilityNCID = RHV
End Property

Public Property Get MedRecNum() As String
    MedRecNum = mvarMedRecNum
End Property
Public Property Let MedRecNum(ByVal RHV As String)
    mvarMedRecNum = RHV
End Property

Public Property Get RadiologyNum() As String
    RadiologyNum = mvarRadiologyNum
End Property
Public Property Let RadiologyNum(ByVal RHV As String)
    mvarRadiologyNum = RHV
End Property




