VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Inpatient_ID  As String
Public Description   As String
Public NCID          As String
Public SnoID     As String
Public ComplaintType As String  '__ one of: NEW or FOLLOW-UP
Public comment       As String

Public Function DeserializeFromRecordset(oRS As ADODB.Recordset)

On Error GoTo ErrHnd

    Inpatient_ID = vbNullString & oRS("INPATIENT_ID")
    Description = vbNullString & oRS("DESCRIPTION")
    NCID = vbNullString & oRS("COMPLAINT_NCID")
    SnoID = vbNullString & oRS("SnoID")
    'MedcinPrefix = vbNullString & oRS("MEDCIN_PREFIX")
    ComplaintType = vbNullString & oRS("COMPLAINT_TYPE")
 
Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cDiagnosis.DeserializeFromRecordset."
    
End Function

Public Function SerializeToRecordset(oRS As ADODB.Recordset)
    oRS("INPATIENT_ID") = Inpatient_ID
    oRS("DESCRIPTION") = Description
    oRS("COMPLAINT_NCID") = NCID
    oRS("SnoID") = SnoID
'    oRS("MEDCIN_PREFIX") = MedcinPrefix
    oRS("COMPLAINT_TYPE") = ComplaintType
End Function


