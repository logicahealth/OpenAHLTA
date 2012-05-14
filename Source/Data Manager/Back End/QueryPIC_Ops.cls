VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryPIC_Ops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDMBE.I_QueryPIC

' Private Class Variables
Private m_strPatientName As String
Private m_strPatientSSN As String
Private m_strPatientUnitNumber As String
Private m_strFileName As String
Private m_datePatientDOB As Date
Private m_enmPICStatus As DMStatus
Private m_colPICDataSets As Collection
Private m_lngDataSets() As Long
Private m_strPatients() As String
Private m_enmPICFormat As DMFormats

Private Property Let I_QueryPIC_AGGPICPatients(ByVal rhs As Variant)
    
    m_strPatients = rhs
    
End Property

Private Property Get I_QueryPIC_AGGPICPatients() As Variant

    I_QueryPIC_AGGPICPatients = m_strPatients
    
End Property

Public Property Let I_QueryPIC_DataSetsInPIC(ByVal rhs As Variant)
    m_lngDataSets = rhs
End Property

Public Property Get I_QueryPIC_DataSetsInPIC() As Variant
    I_QueryPIC_DataSetsInPIC = m_lngDataSets
End Property

Private Sub Class_Initialize()
   Set m_colPICDataSets = New Collection
   Call I_QueryPIC_InitVariables
End Sub

Private Sub Class_Terminate()
   Set m_colPICDataSets = Nothing
   I_QueryPIC_InitVariables
End Sub

Public Sub I_QueryPIC_InitVariables()
   m_strPatientName = ""
   m_strPatientSSN = ""
   m_strPatientUnitNumber = ""
   m_enmPICStatus = 0
End Sub

Public Property Let I_QueryPIC_PatientName(rhs As String)
   m_strPatientName = rhs
End Property

Public Property Get I_QueryPIC_PatientName() As String
   I_QueryPIC_PatientName = m_strPatientName
End Property

Public Property Let I_QueryPIC_PatientSSN(rhs As String)
   m_strPatientSSN = rhs
End Property

Public Property Get I_QueryPIC_PatientSSN() As String
   I_QueryPIC_PatientSSN = m_strPatientSSN
End Property

Public Property Let I_QueryPIC_PatientUnitNumber(rhs As String)
   m_strPatientUnitNumber = rhs
End Property

Public Property Get I_QueryPIC_PatientUnitNumber() As String
   I_QueryPIC_PatientUnitNumber = m_strPatientUnitNumber
End Property

Public Property Let I_QueryPIC_FileName(rhs As String)
   m_strFileName = rhs
End Property

Public Property Get I_QueryPIC_FileName() As String
   I_QueryPIC_FileName = m_strFileName
End Property

Public Property Let I_QueryPIC_PICDataSetExists(Identifier As Variant, rhs As Boolean)
   Dim ClinicalDataArea As String
   Dim TableName As String
   On Error GoTo ErrorTrap
'   If IsNumeric(Identifier) Then
'      If Identifier > 0 And Identifier <= eDM_All Then
'         ClinicalDataArea = GetClinicalDataAreaByEnum(TableName, Identifier)
'         m_colPICDataSets.Add RHS, ClinicalDataArea
'         Exit Property
'      End If
'   Else
'      ClinicalDataArea = GetClinicalDataAreaByName(Identifier, eTab)
      m_colPICDataSets.Add rhs, Identifier
'      Exit Property
'   End If
ErrorTrap:
   Exit Property
End Property

Public Property Get I_QueryPIC_PICDataSetExists(Identifier As Variant) As Boolean
   Dim ClinicalDataArea As String
   Dim eTab As DMDataSets
   On Error GoTo ErrorTrap
   ' assume false
   I_QueryPIC_PICDataSetExists = False
   If IsNumeric(Identifier) Then
      If Identifier > 0 And Identifier <= eDM_All Then
         eTab = Identifier
         ClinicalDataArea = GetClinicalDataAreaByEnum(eTab)
         I_QueryPIC_PICDataSetExists = m_colPICDataSets.Item(ClinicalDataArea)
         Exit Property
      End If
   Else
      ClinicalDataArea = GetClinicalDataAreaByName(Identifier, eTab)
      I_QueryPIC_PICDataSetExists = m_colPICDataSets.Item(Identifier)
      Exit Property
   End If
   Exit Property
ErrorTrap:
      I_QueryPIC_PICDataSetExists = False
End Property

Public Property Let I_QueryPIC_PicStatus(rhs As IDMBE.DMStatus)
   m_enmPICStatus = rhs
End Property

Public Property Get I_QueryPIC_PicStatus() As IDMBE.DMStatus
   I_QueryPIC_PicStatus = m_enmPICStatus
End Property

Public Property Let I_QueryPIC_PicFormat(rhs As IDMBE.DMFormats)
   m_enmPICFormat = rhs
End Property

Public Property Get I_QueryPIC_PicFormat() As IDMBE.DMFormats
   I_QueryPIC_PicFormat = m_enmPICFormat
End Property

'Public Property Let I_QueryPIC_PatientDOB(rhs As Date)
'   m_datePatientDOB = rhs
'End Property
'
'Public Property Get I_QueryPIC_PatientDOB() As Date
'   I_QueryPIC_PatientDOB = m_datePatientDOB
'End Property

