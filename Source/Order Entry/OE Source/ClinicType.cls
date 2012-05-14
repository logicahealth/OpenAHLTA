VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClinicType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClinicCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarClinic_IEN As String 'local copy
Private mvarMEPRS_Code As String 'local copy
Private mvarMEPRS_IEN As String 'local copy
Private mvarClinic_Name As String 'local copy
Private mvarPharmacy_IEN As String 'local copy
Private mvarFormularyGroup_IEN As String 'local copy
Public Property Let Clinic_Name(ByVal vData As String)
    mvarClinic_Name = vData
End Property
Public Property Get Clinic_Name() As String
    Clinic_Name = mvarClinic_Name
End Property
Public Property Let MEPRS_IEN(ByVal vData As String)
    mvarMEPRS_IEN = vData
End Property
Public Property Get MEPRS_IEN() As String
    MEPRS_IEN = mvarMEPRS_IEN
End Property
Public Property Let Pharmacy_IEN(ByVal vData As String)
    mvarPharmacy_IEN = vData
End Property
Public Property Get Pharmacy_IEN() As String
    Pharmacy_IEN = mvarPharmacy_IEN
End Property
Public Property Let FormularyGroup_IEN(ByVal vData As String)
    mvarFormularyGroup_IEN = vData
End Property
Public Property Get FormularyGroup_IEN() As String
    FormularyGroup_IEN = mvarFormularyGroup_IEN
End Property

Public Property Let MEPRS_Code(ByVal vData As String)
    mvarMEPRS_Code = vData
End Property
Public Property Get MEPRS_Code() As String
    MEPRS_Code = mvarMEPRS_Code
End Property
Public Property Let Clinic_IEN(ByVal vData As String)
    mvarClinic_IEN = vData
End Property

Public Property Get Clinic_IEN() As String
    Clinic_IEN = mvarClinic_IEN
End Property



