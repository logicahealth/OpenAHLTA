VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-- frmtransfer history uses a collection of these for display only
'-- give me the wardid, the providerncid, the servicencid,
'   the bed number and the date,

'-- and get out:
'   the specialty name, the location name, the service name, the provider name
'   the bed number and the date
Public TransDate As Date
Public TimeBias As Long
Public BedNumber As String
Public Inpatient_ID As String

Private OldWardID As String
Private oldProviderName As String
Private oldServiceName As String
Private OldSpecialty As String
Private oldServiceNCID As String
Private oldLocation As String
Private OldProviderNCID As String


Public Property Get Specialty() As String
    Specialty = OldSpecialty
End Property
Public Property Get Location() As String
    Location = oldLocation
End Property

Public Property Let WardID(rhs As String)
    OldWardID = rhs
    If Not Val(rhs) <= 0 Then
        oldLocation = gWards(rhs).LocationName
        OldSpecialty = gWards(rhs).SpecialtyName
    End If
End Property

Public Property Let ProviderNCID(rhs As String)
    OldProviderNCID = rhs
    Dim oProv As CHCSII_ClinicClient.Provider
    
    If rhs <> vbNullString Then
        For Each oProv In gColProviders
            If oProv.ClinicianNCID = rhs Then
                oldProviderName = oProv.ProviderName
                Exit For
            End If
        Next oProv
    Else
        oldProviderName = vbNullString
    End If
End Property

Public Property Get ProviderName() As String
    ProviderName = oldProviderName
End Property

Public Property Let ServiceNCID(rhs As String)
    oldServiceNCID = rhs
    If Val(rhs) > 0 Then
        oldServiceName = gSpecialties.GetSpecialty(rhs).SpecName
    End If
End Property

Public Property Get ServiceName() As String
    ServiceName = oldServiceName
End Property

Public Function SaveCurrentStateAsTransfer(opatient As cInpatient, _
                            Optional ByVal DateOfTrans As Date) As Boolean
'__ called from the transfer form and newAdmission (edit) BEFORE the gui to object update.  We save the OLD values here.
'    Dim oTr As cTransfer
'    Set oTr = New cTransfer
    On Error GoTo ErrHandler

    Inpatient_ID = opatient.ID
    BedNumber = opatient.AssignedBedNumber
    If Not opatient.Bed Is Nothing Then
        If Not opatient.Bed.Ward Is Nothing Then
            If Val(opatient.Bed.Ward.WardID) > 0 Then
                WardID = opatient.Bed.Ward.WardID
            Else
                WardID = opatient.Ward_ID
            End If
        Else
            If Val(opatient.Ward_ID) > 0 Then
                WardID = opatient.Ward_ID
            End If
        End If
    End If
    If CDbl(DateOfTrans) = 0 Then
        TransDate = Format$(gobjComm.GetDateTime, "DD-MMM-YYYY hh:nn:ss")  '-get utc
    Else
        TransDate = gobjComm.LocalToGmt(DateOfTrans)    '   CDate(Format$(dtDate, "DD-MMM-YYYY") & " " & Format$(dtTime, "hh:MM"))
    End If
    TimeBias = gobjComm.GetTotalTimeBias '105144
    ProviderNCID = opatient.AttendingProviderNCID
    ServiceNCID = opatient.PrimaryServiceNCID

    Call Save

    Exit Function
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "cTransfer.SaveCurrentStateAsTransfer", "error in admission", vbCritical
End Function

Public Function Save() As Boolean
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim oRS As ADODB.Recordset
    Set oRS = oDal.GetInpatientData(edti_Transfer, Inpatient_ID)
    oRS.AddNew
    Call SerializeToRecordset(oRS)
    oRS.Update
    oRS.Close
End Function

Public Function DeserializeFromRecordset(oRS As ADODB.Recordset)
On Error Resume Next
    Inpatient_ID = vbNullString & oRS("Inpatient_ID")
    TransDate = vbNullString & oRS("TRANSFER_DATE")
    TimeBias = Val(vbNullString & oRS("Timebias"))
    If Val(vbNullString & oRS("WARD_NCID")) > 0 Then
        WardID = vbNullString & oRS("WARD_NCID")
    End If
    If Val(vbNullString & oRS("ASSIGNED_BED_NUMBER")) > 0 Then
    BedNumber = vbNullString & oRS("ASSIGNED_BED_NUMBER")
    End If
    ProviderNCID = vbNullString & oRS("ATTENDING_PROVIDER_NCID")
    ServiceNCID = vbNullString & oRS("PRIMARY_SERVICE_NCID")
End Function


Public Function SerializeToRecordset(oRS As ADODB.Recordset)
    oRS("Inpatient_ID") = Inpatient_ID
    oRS("TRANSFER_DATE") = TransDate '  CDate(Format$(dtDate, "DD-MMM-YYYY") & " " & Format$(dtTime, "hh:MM"))
    oRS("timebias") = TimeBias
    oRS("ASSIGNED_BED_NUMBER") = Val(BedNumber)
    oRS("WARD_NCID") = Val(OldWardID)
    oRS("ATTENDING_PROVIDER_NCID") = Val(OldProviderNCID)
    oRS("PRIMARY_SERVICE_NCID") = Val(oldServiceNCID)
End Function
