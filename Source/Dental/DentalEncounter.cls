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
Attribute VB_Name = "DentalEncounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbNewDentalEncounter As Boolean

Private mstrPSRlowerLeft As String
Private mstrPSRlowerCenter As String
Private mstrPSRlowerRight As String
            
Private mstrPSRupperLeft As String
Private mstrPSRupperCenter As String
Private mstrPSRupperRight As String

Private mlngDentalEncounterID As Long
Private mlngPatUnitId As Long
Private mlngMedicalEncounterID As Long
Private mdteEncDate As Date

Private mlngEligibilityFlag As Long
Private mdteEligibilityDate As Date

Private mobjSBE As SBE
Private mobjDRC As DRC
Private svgCommObj As srvSVGComm.clsSVGComm

Implements IntrfcCHCSII_Dental.IDental_Enc

Private Property Get IDental_Enc_DentalEncounterId() As Long
    IDental_Enc_DentalEncounterId = mlngDentalEncounterID
End Property

Private Property Get IDental_Enc_DRC() As IntrfcCHCSII_Dental.IDENTAL_DRC
    Set IDental_Enc_DRC = mobjDRC
End Property

Private Property Get IDental_Enc_EligibilityDate() As Date

End Property

Private Property Get IDental_Enc_EligibilityFlag() As Boolean

End Property

Private Property Get IDental_Enc_EncounterDate() As Date
    IDental_Enc_EncounterDate = mdteEncDate
End Property

Private Function IDental_Enc_LoadDentalEncounter(ByVal lngPatientID As Long, ByVal lngEncounterId As Long, ByVal lngFacilityNCID As Long, ByVal dteApptDate As Date) As Boolean
    
    Dim objDAS As ICHCSII_DataAccess.ICHCSII_DAS
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim svgCommFactory As srvSVGComm.SingletonFactory

    Dim objRS As ADODB.Recordset
    Dim intEncLvlDRC As Integer
    Dim intPatLvlDRC As Integer
    Dim dteEncLvlLstExmDte As Date
    Dim dtePatLvlLstExmDte As Date
    Dim dteEncLvlLstChgDte As Date
    Dim dtePatLvlLstChgDte As Date
    Dim blnReturn As Boolean
    Dim strLastExamDate As String
    Dim strLastChangeDate As String
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = lngPatientID
    mlngMedicalEncounterID = lngEncounterId
    mdteEncDate = dteApptDate
    
    mbNewDentalEncounter = False
    
    'Set objDAS = New CHCSII_DataAccess.CHCSII_DASsP
    Set objDAS = GetDasObj
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    Set mobjSBE = New SBE
    Set mobjDRC = New DRC
    
    Set svgCommFactory = New srvSVGComm.SingletonFactory  'hwcc 2/26/2002
    If Not (svgCommFactory Is Nothing) Then Set svgCommObj = svgCommFactory.Singleton
    
    objCmd.SPName = "DENT_PARENT_OPEN_DENT_ENC"
    'We must add paramaters in the same order as the appear in the Stored Proc Signature
    Call objCmd.AddParam(False, False, "textString", lngPatientID)
    Call objCmd.AddParam(False, False, "textString", lngEncounterId)
    Call objCmd.AddParam(False, False, "textString", lngFacilityNCID)
    Call objCmd.AddParam(False, False, "date", dteApptDate)
    
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")
    Set objRS = objDAS.OpenRecordsetSP(objCmd)
    
    
    If Not objRS.RecordCount = 0 Then
    ' found an existing encounter of patient
    'mLastDentalDate = objRS(0)
        With objRS
            mlngDentalEncounterID = .Fields("DENTAL_ENCOUNTER_ID")
            intEncLvlDRC = .Fields("DRC") & ""
            strLastExamDate = .Fields("DRC_LastExamDate") & ""
            strLastChangeDate = .Fields("DRC_LastChangeDate") & ""
            mbNewDentalEncounter = .Fields("NEW_ENCOUNTER")
            intPatLvlDRC = Val(.Fields("DENTAL_READINESS_DRC") & "")
            dtePatLvlLstChgDte = .Fields("DENTAL_READINESS_DT") & ""
            dtePatLvlLstExmDte = .Fields("LAST_DENTAL_EXAM_DT") & ""
            mlngEligibilityFlag = .Fields("ELIGIBILITYFLAG") & ""
            mdteEligibilityDate = .Fields("ELIGIBILITYDATE") & ""
            
            mobjSBE.Set_SBE_Info Val(.Fields("SBE") & ""), .Fields("SBE_COMMENTS") & ""
        End With
        
        'Set the PSR info. This is used by Graphic Charting, and Dental Parent to generate the
        'RTF note
        With svgCommObj.mObjPSR
           .lowerLeft = objRS("PSR_LOWER_LEFT") & ""
           .lowerCenter = objRS("PSR_LOWER_CENTER") & ""
           .lowerRight = objRS("PSR_LOWER_RIGHT") & ""
                       
           .upperLeft = objRS.Fields("PSR_UPPER_LEFT") & ""
           .upperCenter = objRS.Fields("PSR_UPPER_CENTER") & ""
           .upperRight = objRS.Fields("PSR_UPPER_RIGHT") & ""
           .iEncounterID = lngEncounterId
           .iPatientID = lngPatientID
           .iEncounterID = mlngDentalEncounterID
        End With
        
        If strLastExamDate <> "12/31/9999" Or strLastExamDate <> Empty Then
            dteEncLvlLstExmDte = CDate(strLastExamDate)
        Else
            'problem with the value in db, use patient level value
            dteEncLvlLstExmDte = dtePatLvlLstExmDte
        End If
        
        If strLastChangeDate <> "12/31/9999" Or strLastExamDate <> Empty Then
            dteEncLvlLstChgDte = strLastChangeDate
        Else
            dteEncLvlLstChgDte = dtePatLvlLstChgDte
        End If
        
        'Set the DRC info
        mobjDRC.SetEncounterLevelDRCInfo intEncLvlDRC, dteEncLvlLstExmDte, dteEncLvlLstChgDte
        mobjDRC.SetPatLevelDRCInfo intPatLvlDRC, dtePatLvlLstExmDte, dtePatLvlLstChgDte
        blnReturn = True
        
    Else
        blnReturn = False
        
    End If
    
FuncExit:
    Set objRS = Nothing
    Set objCmd = Nothing
    Set objDAS = Nothing
    Set svgCommFactory = Nothing
    
    IDental_Enc_LoadDentalEncounter = blnReturn
    
    Exit Function
    
ErrHandler:
    'SCR 38333 RWalker 11/07/2003 Full dump all data for Datatype Mismatch error
    If Err.Number = 13 Then 'Datatype Mismatch
        Dim sErrTxt As String
        sErrTxt = "Dental.LoadDentalEncounter Error(" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf
        sErrTxt = sErrTxt & vbTab & vbTab & "Data dump:" & vbCrLf & vbCrLf
        sErrTxt = sErrTxt & "In Parms: PatId=" & lngPatientID & " EncId=" & lngEncounterId & " FacId=" & lngFacilityNCID & " ApptDate=" & dteApptDate & vbCrLf
        sErrTxt = sErrTxt & " Db data: DtlEncId=" & objRS.Fields("DENTAL_ENCOUNTER_ID") & " NewEncId=" & objRS.Fields("NEW_ENCOUNTER") & vbCrLf
        sErrTxt = sErrTxt & "DRC=" & objRS.Fields("DRC") & " DRCLastExamDt=" & objRS.Fields("DRC_LastExamDate") & " DRCLastChgDt=" & objRS.Fields("DRC_LastChangeDate") & vbCrLf
        sErrTxt = sErrTxt & "DtlRdyDRC=" & objRS.Fields("DENTAL_READINESS_DRC") & "DtlRdyDt=" & objRS.Fields("DENTAL_READINESS_DT") & vbCrLf
        sErrTxt = sErrTxt & "LstDtlExamDt=" & objRS.Fields("LAST_DENTAL_EXAM_DT") & " EligFlag=" & objRS.Fields("ELIGIBILITYFLAG") & " EligDt=" & objRS.Fields("ELIGIBILITYDATE") & vbCrLf
        sErrTxt = sErrTxt & "SBE=" & objRS.Fields("SBE") & " SBE Cmts=" & objRS.Fields("SBE_COMMENTS") & vbCrLf
        sErrTxt = sErrTxt & "PSRLwrLft=" & objRS.Fields("PSR_LOWER_LEFT") & " PSRLwrCnt=" & objRS.Fields("PSR_LOWER_CENTER") & " PSRLwrRgt=" & objRS.Fields("PSR_LOWER_RIGHT") & vbCrLf
        sErrTxt = sErrTxt & "PSRUprLft=" & objRS.Fields("PSR_UPPER_LEFT") & " PSRUprCnt=" & objRS.Fields("PSR_UPPER_CENTER") & " PSRUprRgt=" & objRS.Fields("PSR_UPPER_RIGHT") & vbCrLf
#If Not DEBUGON Then
    Dim DL As cDebugLog
    Set DL = GetObject("", "dlog.cdebuglog")
    DL.debugStart
#End If
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog App.EXEName, "DentalEncounter.cls", "Dental.LoadDentalEncounter", sErrTxt, DR
#If Not DEBUGON Then
    If Not DL Is Nothing Then DL.debugStop
    If Not DL Is Nothing Then Set DL = Nothing
#End If
        sErrTxt = vbTab & "Patient SBE information could not be retrieved." & vbCrLf & vbCrLf
        sErrTxt = sErrTxt & "Please check this patient's SBE status, and document in the encounter note."
        MsgBxARMd sErrTxt, vbSystemModal + vbInformation + vbOKOnly
        Err.Clear
        Resume Next
    Else
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.LoadDentalEncounter", App.Title, vbExclamation)
    End If

    blnReturn = False
    GoTo FuncExit
    Resume
End Function

Private Property Get IDental_Enc_MedicalEncounterId() As Long
    IDental_Enc_MedicalEncounterId = mlngMedicalEncounterID
End Property

Private Property Get IDental_Enc_NewEncounter() As Boolean
    IDental_Enc_NewEncounter = mbNewDentalEncounter
End Property

Private Property Get IDental_Enc_PatientId() As Long
    IDental_Enc_PatientId = mlngPatUnitId
End Property

Private Property Get IDental_Enc_SBE() As IntrfcCHCSII_Dental.IDENTAL_SBE
    Set IDental_Enc_SBE = mobjSBE
End Property

'*************************************************************************************
'Sub:SetDentalEncounter
'Author:KDunne
'Date:9/9/02
'Purpose: Sub is called when the stored procs are not available and encounter data is
'         loaded from inline sql.
'*************************************************************************************
Public Sub SetDentalEncounter(ByVal lngDentEncId As Long, ByVal lngPatUnitId As Long, _
            ByVal lngMedEncId As Long, ByVal lngFacId As Long, ByVal dteApptDte As Date, _
            ByVal blnNewEnc As Boolean, _
            ByVal intEncDRC As Integer, ByVal dteEncLstExmDte As Date, _
            ByVal dteEncLstChgDate As Date, ByVal intPatDRC As Integer, ByVal dtePatLstExmDte As Date, _
            ByVal dtePatLstChgDate As Date, ByVal intSBE As Integer, ByVal strComments As String)

 mbNewDentalEncounter = blnNewEnc

 mlngDentalEncounterID = lngDentEncId
 mlngPatUnitId = lngPatUnitId
 mlngMedicalEncounterID = lngMedEncId
 mdteEncDate = dteApptDte

Set mobjDRC = New DRC
Set mobjSBE = New SBE

mobjDRC.SetEncounterLevelDRCInfo intEncDRC, dteEncLstExmDte, dteEncLstChgDate
mobjDRC.SetPatLevelDRCInfo intPatDRC, dtePatLstExmDte, dtePatLstChgDate

mobjSBE.Set_SBE_Info intSBE, strComments

End Sub
Private Sub Class_Terminate()
    Set mobjSBE = Nothing
    Set mobjDRC = Nothing
    Set svgCommObj = Nothing
End Sub
