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
Attribute VB_Name = "clsTrNSEEnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsISection

Private mlngDentalEncounterId As Long
Private mlngPatUnitId As Long
Private mlngDirtyFlag As DirtyFlagEnum
Private mlngId As Long

Private mstrSuppProc As String
Private mstrComments As String
Private mblnUnscheduledVisit As Boolean
Private mobjSOAP As clsExamSOAP


'Public Enum TreatEndo_SupptProcsEnum
'    Not_Addressed
'    AlteredCastImpression
'    SalivaryGlandMilking
'    Intraoral
'    ToluidienBlueTissueStain
'    Aspiration
'    BorderMolding
'    CottonRolls
'    DeliveryProsthesis
'    DentinBonding
'    DirectPattern
'    DrainPlacement
'    FinalImpression
'    FollowUp
'    FrameWork
'    GuttaPerchaCone
'    HemstaticAgent
'    InterimProsthesisFab
'    InterimProsthesisPlace
'    InterimRestoration
'    IntracanalIrrigation
'    IntracanalMedicament
'    Methylene
'    OcclusalReduction
'    PostOpFollowUp
'    PDLRemoval
'    RemoveCrown
'    RetractionCord
'    RPDMouthPrep
'    RuberDam
'    SauterIndex
'    ShadeSelection
'    Soak
'    ToothPrep
'    ToothSelection
'    Ultrasonics
'    Wax
'
'End Enum
Private Const const_ModuleName = "clsTrNSEEnc"

Public Property Get Supporting_Procedures() As String
    Supporting_Procedures = mstrSuppProc
End Property
Public Property Let Supporting_Procedures(ByVal strValue As String)
    mstrSuppProc = strValue
    UpdateDirtyFlag
End Property
Public Property Get Comments() As String
    Comments = mstrComments
End Property
Public Property Let Comments(ByVal strValue As String)
    mstrComments = strValue
    UpdateDirtyFlag
End Property

Public Property Get Unscheduled_Visit() As Boolean
    Unscheduled_Visit = mblnUnscheduledVisit
End Property
Public Property Let Unscheduled_Visit(ByVal blnValue As Boolean)
    mblnUnscheduledVisit = blnValue
    UpdateDirtyFlag
End Property
Public Property Set SOAP(ByVal objSOAP As clsExamSOAP)
    Set mobjSOAP = objSOAP
End Property
Public Property Get SOAP()
    Set SOAP = mobjSOAP
End Property
Public Property Get NSE_Endo_Encounter_ID() As Long
    NSE_Endo_Encounter_ID = mlngId
End Property

Private Sub Class_Initialize()
    mstrSuppProc = gcSupportingProcs
End Sub

Private Sub Class_Terminate()
    Set mobjSOAP = Nothing
End Sub

Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim strTemp As String
    
    strNote = Empty
    
    If mblnUnscheduledVisit Then
        strNote = " Unscheduled visit:" & ";"
    End If
        
    If Trim(mstrComments) <> "" Then
        strNote = strNote & " Comments = " & mstrComments & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrSuppProc)
    
    If Trim(strTemp) <> "" Then
        strNote = strNote & strTemp & "; "
    End If
    
    If strNote <> "" Then
        'Section has data. Prepend a title and return the data.
        strNote = "Endodontics Encounter: " & strNote
    End If
    clsISection_GetNote = strNote
End Function

Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim strTemp As String
    
    strNote = Empty
    
    If mblnUnscheduledVisit Then
        strNote = " Unscheduled visit:" & ";"
    End If
        
    If Trim(mstrComments) <> "" Then
        strNote = strNote & " Comments = " & mstrComments & ";"
    End If
    
    strTemp = Get_MultiSelect_DisplayString(mstrSuppProc)
    
    If Trim(strTemp) <> "" Then
        strNote = strNote & strTemp & "; "
    End If
    
    If strNote <> "" Then
        'Section has data. Prepend a title and return the data.
        strNote = crtfStartControlName & "Endodontics Encounter: " & _
            crtfEndControlName & strNote
    End If
    clsISection_GetRTFNote = strNote
End Function

Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set objRS = DB_LoadNSEEndoEnco(mlngDentalEncounterId)
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            'reset property values here
            mstrComments = Empty
            mblnUnscheduledVisit = False
            mstrSuppProc = gcSupportingProcs
        Else
            mlngDirtyFlag = dfUpdate
            'Set property values here
            mstrComments = .Fields("COMMENTS") & ""
            mblnUnscheduledVisit = CBool(Val(.Fields("UNSCHEDULED_VISIT") & ""))
            mlngId = Val(.Fields("EE_ENC_ID") & "")
            mstrSuppProc = .Fields("SUPPORT_PROC") & ""
        End If
    End With
    
FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".LoadSection()", "", vbCritical

   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    
    On Error GoTo ErrHandler
    
    strSql = Empty
    Select Case mlngDirtyFlag
        Case dfNoAction
            blnReturn = True
            GoTo FuncExit
        Case dfAdd
        
            mlngId = goconn.GetNextID(esTEMPL_TREAT_NSE_ENCOUNTER_SEQ)
'            strSql = "SELECT TEMPL_TREAT_NSE_ENCOUNTER_seq.NEXTVAL From Dual"
'            Set objRS = goconn.OpenRecordset(strSql)
'            mlngId = objRS.Fields("NEXTVAL")
            
            strSql = "INSERT INTO TEMPL_TREAT_NSE_ENCOUNTER " & _
                     "(EE_ENC_ID, DENTAL_ENCOUNTER_ID, " & _
                     " UNSCHEDULED_VISIT, COMMENTS, SUPPORT_PROC) " & _
                     "VALUES (" & mlngId & ", " & mlngDentalEncounterId & ", " & _
                     CInt(mblnUnscheduledVisit) & ", '" & RQ(mstrComments) & "','" & mstrSuppProc & "')"

        Case dfUpdate
            strSql = "UPDATE TEMPL_TREAT_NSE_ENCOUNTER " & _
                     " SET DENTAL_ENCOUNTER_ID = " & mlngDentalEncounterId & ", " & _
                     " COMMENTS = '" & Scrub_Apostrophe(mstrComments) & "', UNSCHEDULED_VISIT = " & _
                     CInt(mblnUnscheduledVisit) & ", SUPPORT_PROC = '" & mstrSuppProc & _
                     "' WHERE EE_ENC_ID = " & mlngId
        Case dfdelete
            strSql = "DELETE FROM TEMPL_TREAT_NSE_ENCOUNTER WHERE " & _
                     " EE_ENC_ID = " & mlngId
    End Select
    
    If strSql <> "" Then
        goconn.ExecuteSQL (strSql)
    End If
    mlngDirtyFlag = dfNoAction
    
    'If there's a collection, loop through it here.
    
FuncExit:
    Exit Function
ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".SaveSection()", "", vbCritical
    GoTo FuncExit
End Function

'****************************************************
'Sub: UpdateDirtyFlag
'Author: KDunne
'Date: 11/5/2001
'Purpose: Updates module level dirty flag field
'****************************************************
Private Sub UpdateDirtyFlag()

    If mlngDirtyFlag = dfNoAction Then
        mlngDirtyFlag = dfUpdate
    End If
    
End Sub

Private Function DB_LoadNSEEndoEnco(ByVal lngDentalEncId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
    strSql = "Select EE_ENC_ID, DENTAL_ENCOUNTER_ID, Unscheduled_Visit , Comments, SUPPORT_PROC" & _
        " From TEMPL_TREAT_NSE_ENCOUNTER WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncId
        
    Set DB_LoadNSEEndoEnco = goconn.OpenRecordset(strSql)
    
    Exit Function
ErrHandler:
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_LoadNSEEndoEnco()", "", vbCritical

End Function
