VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExProsOccl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Additional Public constants are defined in the modDentureAssessment.bas file.
'Only a very few so there was no need to create its own file.

Implements clsISection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum
Public colMaterialUsed As New Collection

Private mstrCentricRelation As String
Private mstrLeavesAtContact As String
Private mstrPointOfContact As String
Private mstrMaterialUsed As String
Private mstrCarrierType As String
Private msinA_P_Slide As Single
Private msinLateralSlide As Single
Private msinRtLateralGuidance As Single
Private msinLtLateralGuidance As Single
Private mstrProtrusiveGuidance As String
Private mstrRtAnteriorCoupling As String
Private mstrLtAnteriorCoupling As String
Private mstrCRreproducibility As String

 Const gcOcclusalMaterials As String = "Polyvinylsiloxane|" _
    & "Polyether|Zinc Oxide|" _
    & "Plaster|Modeling Plastic|Wax|" _
    & "Acrylic Resin|"

Public Property Let CRreproducibility(ByVal vData As String)
    mstrCRreproducibility = vData
    UpdateDirtyFlag
End Property

Public Property Get CRreproducibility() As String
    CRreproducibility = mstrCRreproducibility
End Property

Public Property Let RtAnteriorCoupling(ByVal vData As String)
    mstrRtAnteriorCoupling = vData
    UpdateDirtyFlag
End Property
Public Property Get RtAnteriorCoupling() As String
    RtAnteriorCoupling = mstrRtAnteriorCoupling
End Property
Public Property Let LtAnteriorCoupling(ByVal vData As String)
    mstrLtAnteriorCoupling = vData
    UpdateDirtyFlag
End Property
Public Property Get LtAnteriorCoupling() As String
    LtAnteriorCoupling = mstrLtAnteriorCoupling
End Property
Public Property Let ProtrusiveGuidance(ByVal vData As String)
    mstrProtrusiveGuidance = vData
    UpdateDirtyFlag
End Property
Public Property Get ProtrusiveGuidance() As String
    ProtrusiveGuidance = mstrProtrusiveGuidance
End Property
Public Property Let LtLateralGuidance(ByVal vData As Single)
    msinLtLateralGuidance = vData
    UpdateDirtyFlag
End Property
Public Property Get LtLateralGuidance() As Single
    LtLateralGuidance = msinLtLateralGuidance
End Property
Public Property Let RtLateralGuidance(ByVal vData As Single)
    msinRtLateralGuidance = vData
    UpdateDirtyFlag
End Property
Public Property Get RtLateralGuidance() As Single
    RtLateralGuidance = msinRtLateralGuidance
End Property
Public Property Let LateralSlide(ByVal vData As Single)
    msinLateralSlide = vData
    UpdateDirtyFlag
End Property
Public Property Get LateralSlide() As Single
    LateralSlide = msinLateralSlide
End Property
Public Property Let A_P_Slide(ByVal vData As Single)
    msinA_P_Slide = vData
    UpdateDirtyFlag
End Property
Public Property Get A_P_Slide() As Single
    A_P_Slide = msinA_P_Slide
End Property
Public Property Let CarrierType(ByVal vData As String)
    mstrCarrierType = vData
    UpdateDirtyFlag
End Property
Public Property Get CarrierType() As String
    CarrierType = mstrCarrierType
End Property
Public Property Let MaterialUsed(ByVal vData As String)
    mstrMaterialUsed = vData
    UpdateDirtyFlag
End Property
Public Property Get MaterialUsed() As String
    If mstrMaterialUsed = "" Then
        mstrMaterialUsed = gcOcclusalMaterials
    End If
    MaterialUsed = mstrMaterialUsed
End Property
Public Property Let PointOfContact(ByVal vData As String)
    mstrPointOfContact = vData
    UpdateDirtyFlag
End Property
Public Property Get PointOfContact() As String
    PointOfContact = mstrPointOfContact
End Property
Public Property Let LeavesAtContact(ByVal vData As String)
    mstrLeavesAtContact = vData
    UpdateDirtyFlag
End Property
Public Property Get LeavesAtContact() As String
    LeavesAtContact = mstrLeavesAtContact
End Property
Public Property Let CentricRelation(ByVal vData As String)
    mstrCentricRelation = vData
    UpdateDirtyFlag
End Property
Public Property Get CentricRelation() As String
    CentricRelation = mstrCentricRelation
End Property
Private Property Get clsISection_DentalEncounterID() As Long
    
    clsISection_DentalEncounterID = mvarDentalEncounterID

End Property


Private Property Get clsISection_PatientID() As Long
    
    clsISection_PatientID = mvarPatientID

End Property



Private Function clsISection_GetRTFNote() As String
 Dim sNote As String
 Dim sTxt As String
 Dim strFindNote As String
 
 Const sPROC_NAME = "clsExProsOccl.clsISection_GetRTFNote"
 On Error GoTo errTrap
    
    If mstrCentricRelation = "" And _
    mstrLeavesAtContact = "" And _
    mstrPointOfContact = "" And _
    mstrMaterialUsed = "" And _
    mstrCarrierType = "" And _
    msinA_P_Slide = -1 And _
    msinLateralSlide = -1 And _
    msinRtLateralGuidance = -1 And _
    msinLtLateralGuidance = -1 And _
    mstrProtrusiveGuidance = "" And _
    mstrRtAnteriorCoupling = "" And _
    mstrLtAnteriorCoupling = "" And _
    mstrCRreproducibility = "" Then
    clsISection_GetRTFNote = ""
    Exit Function
    End If

    strFindNote = Empty
       
    If Trim(mstrCentricRelation) <> "" Then strFindNote = strFindNote & "Centric Relation Record: " & mstrCentricRelation & "; "
    If Trim(mstrPointOfContact) <> "" Then strFindNote = strFindNote & "1st Point of Contact: " & mstrPointOfContact & "; "
    If Trim(mstrLeavesAtContact) <> "" Then strFindNote = strFindNote & "Leaves at 1st Contact: " & mstrLeavesAtContact & "; "
    If msinA_P_Slide <> -1 Then strFindNote = strFindNote & "A-P Slide: " & CStr(msinA_P_Slide) & " mm" & "; "
    If msinLateralSlide <> -1 Then strFindNote = strFindNote & "Lateral Slide: " & CStr(msinLateralSlide) & " mm" & "; "
    If msinRtLateralGuidance <> -1 Then strFindNote = strFindNote & "Rt. Lateral Guidance: " & CStr(msinRtLateralGuidance) & " mm" & "; "
    If msinLtLateralGuidance <> -1 Then strFindNote = strFindNote & "Lt. Lateral Guidance: " & CStr(msinLtLateralGuidance) & " mm" & "; "
    If Trim(mstrProtrusiveGuidance) <> "" Then strFindNote = strFindNote & "Protrusive Guidance: " & mstrProtrusiveGuidance & "; "
    If Trim(mstrRtAnteriorCoupling) <> "" Then strFindNote = strFindNote & "Rt. Anterior Coupling: " & mstrRtAnteriorCoupling & "; "
    If Trim(mstrLtAnteriorCoupling) <> "" Then strFindNote = strFindNote & "Lt. Anterior Coupling: " & mstrLtAnteriorCoupling & "; "
    If Trim(mstrCRreproducibility) <> "" Then strFindNote = strFindNote & "C/R Reproducibility: " & mstrCRreproducibility & "; "
        
    If Trim(strFindNote) <> "" Then
        'Remove the last ";
        strFindNote = Left(strFindNote, Len(strFindNote) - Len("; "))
        sNote = crtfStartControlName & "Findings: " & strFindNote & crtfEndControlName
    End If
    
    
    sTxt = MaterialUsedDisplay()
    If sTxt <> "" Or mstrCarrierType <> "" Then
        sNote = sNote + crtfStartControlName & "Material Used: "
        If sTxt <> "" Then
            sNote = sNote + sTxt + "; "
        End If
        If mstrCarrierType <> "" Then
            If mstrCarrierType = "W" Then
                sNote = sNote + "With Carrier; "
            Else
                sNote = sNote + "Without Carrier; "
            End If
        End If
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote + crtfEndControlName
    End If
    
    
clsISection_GetRTFNote = sNote
    
Exit Function
errTrap:
    clsISection_GetRTFNote = "Error in generating note for ExamEval-Pros Occlusal Analysis section."
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function
'*******************************************************************
'*******************************************************************
'Function: MaterialUsedDisplay
'Author: Sean McAvoy
'Date: 1/20/02
'Returns: String
'Purpose: To get the user viewable display string for Material Used.
'*******************************************************************
Public Function MaterialUsedDisplay() As String

Dim oItems As New CMultiListItems
    oItems.ListType = Standard
    oItems.LoadFromValueString mstrMaterialUsed
    MaterialUsedDisplay = oItems.Value_String_Display

End Function

'*******************************************************************
'Function: GetNote
'Author: Sean McAvoy
'Date: 12/19/10
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

 Dim sNote As String
 Dim sTxt As String
 
 Const sPROC_NAME = "clsExProsOccl.clsISection_GetNote"
 On Error GoTo errTrap
    
    If mstrCentricRelation = "" And _
    mstrLeavesAtContact = "" And _
    mstrPointOfContact = "" And _
    mstrMaterialUsed = "" And _
    mstrCarrierType = "" And _
    msinA_P_Slide = -1 And _
    msinLateralSlide = -1 And _
    msinRtLateralGuidance = -1 And _
    msinLtLateralGuidance = -1 And _
    mstrProtrusiveGuidance = "" And _
    mstrRtAnteriorCoupling = "" And _
    mstrLtAnteriorCoupling = "" And _
    mstrCRreproducibility = "" Then
    clsISection_GetNote = ""
    Exit Function
    End If

    sNote = sNote + "Exam Evaluation - Prosthetics - Occlusal Analysis:" + vbCrLf
    If mstrCarrierType <> "" Then
        If mstrCarrierType = "W" Then
            sTxt = " With Carrier"
        Else
            sTxt = " Without Carrier"
        End If
    End If
    If mstrMaterialUsed <> "" Then sNote = sNote + vbTab + "Material Used" + sTxt + ": " + Replace(Me.MaterialUsed, ",", ", ") + vbCrLf
    If mstrCentricRelation <> "" Then sNote = sNote + vbTab + "Centric Relation: " + Me.CentricRelation + vbCrLf
    If mstrPointOfContact <> "" Then sNote = sNote + vbTab + "1st Point of Contact: " + Me.PointOfContact + vbCrLf
    If mstrLeavesAtContact <> "" Then sNote = sNote + vbTab + "Leaves at 1st Contact: " + Me.LeavesAtContact + vbCrLf
    If msinA_P_Slide <> -1 Then sNote = sNote + vbTab + "A-P Slide: " + CStr(Me.A_P_Slide) + " mm" + vbCrLf
    If msinLateralSlide <> -1 Then sNote = sNote + vbTab + "Lateral Slide: " + CStr(Me.LateralSlide) + " mm" + vbCrLf
    If msinRtLateralGuidance <> -1 Then sNote = sNote + vbTab + "Rt. Lateral Guidance: " + CStr(Me.RtLateralGuidance) + " mm" + vbCrLf
    If msinLtLateralGuidance <> -1 Then sNote = sNote + vbTab + "Lt. Lateral Guidance: " + CStr(Me.LtLateralGuidance) + " mm" + vbCrLf
    If mstrProtrusiveGuidance <> "" Then sNote = sNote + vbTab + "Protrusive Guidance: " + Me.ProtrusiveGuidance + vbCrLf
    If mstrRtAnteriorCoupling <> "" Then sNote = sNote + vbTab + "Rt. Anterior Coupling: " + Me.RtAnteriorCoupling + vbCrLf
    If mstrLtAnteriorCoupling <> "" Then sNote = sNote + vbTab + "Lt. Anterior Coupling: " + Me.LtAnteriorCoupling + vbCrLf
    If mstrCRreproducibility <> "" Then sNote = sNote + vbTab + "C/R Reproducibility: " + Me.CRreproducibility + vbCrLf

    sNote = sNote + vbCrLf + vbCrLf

clsISection_GetNote = sNote
    
    
Exit Function
errTrap:
    clsISection_GetNote = "Error in generating note for ExamEval-Pros Occlusal Analysis section."
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function


'*******************************************************************
'Function: SaveSection
'Author: Sean McAvoy
'Date: 12/19/01
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String
Dim mstrLeavesAtContact_TMP As String
Dim mstrPointOfContact_TMP As String

Const sPROC_NAME As String = "clsExProsOccl.SaveSection"

On Error GoTo ErrHnd
    
  
    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    End If
    If mvarDentalEncounterID = 0 Then
        clsISection_SaveSection = False
        Exit Function
    End If
    'Remove ' from the free text properties
    mstrLeavesAtContact_TMP = RQ(mstrLeavesAtContact)
    mstrPointOfContact_TMP = RQ(mstrPointOfContact)
        
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_EXAM_PROS_OCCLUSAL (" _
            & " DENTAL_ENCOUNTER_ID, CENTRICRELATION, LEAVESATCONTACT, POINTOFCONTACT, MATERIALUSED, CARRIERTYPE, A_P_SLIDE," _
            & " LATERALSLIDE, RTLATGUIDANCE, LTLATGUIDANCE, PROTGUIDANCE, RTANTERIORCUPLING, LTANTERIORCUPLING, CRREPRODUCIBILITY) " _
            & " VALUES" _
            & " (" & CStr(mvarDentalEncounterID) & ", " _
            & " '" & mstrCentricRelation & "', " _
            & " '" & mstrLeavesAtContact_TMP & "', " _
            & " '" & mstrPointOfContact_TMP & "', " _
            & " '" & mstrMaterialUsed & "', " _
            & " '" & mstrCarrierType & "', " _
            & CStr(msinA_P_Slide) & ", " _
            & CStr(msinLateralSlide) & ", " _
            & CStr(msinRtLateralGuidance) & ", " _
            & CStr(msinLtLateralGuidance) & ", " _
            & " '" & mstrProtrusiveGuidance & "', " _
            & " '" & mstrRtAnteriorCoupling & "', " _
            & " '" & mstrLtAnteriorCoupling & "', " _
            & " '" & mstrCRreproducibility & "' )"
            
    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
            sSQL = "UPDATE TEMPL_EXAM_PROS_OCCLUSAL SET " + _
            "CENTRICRELATION = '" + mstrCentricRelation + "', " + _
            "LEAVESATCONTACT = '" + mstrLeavesAtContact_TMP + "', " + _
            "POINTOFCONTACT = '" + mstrPointOfContact_TMP + "', " + _
            "MATERIALUSED = '" + mstrMaterialUsed + "', " + _
            "CARRIERTYPE = '" + mstrCarrierType + "', " + _
            "A_P_SLIDE = " + CStr(msinA_P_Slide) + ", " + _
            "LATERALSLIDE = " + CStr(msinLateralSlide) + ", " + _
            "RTLATGUIDANCE = " + CStr(msinRtLateralGuidance) + ", " + _
            "LTLATGUIDANCE = " + CStr(msinLtLateralGuidance) + ", " + _
            "PROTGUIDANCE = '" + mstrProtrusiveGuidance + "', " + _
            "RTANTERIORCUPLING = '" + mstrRtAnteriorCoupling + "', " + _
            "LTANTERIORCUPLING = '" + mstrLtAnteriorCoupling + "', " + _
            "CRREPRODUCIBILITY = '" + mstrCRreproducibility + "'" + _
            " Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
            
    ElseIf dfDirtyFlag = dfdelete Then
'        sSQL = "DELETE FROM TEMPL_EXAM_PROS_OCCLUSAL " + _
'               "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
    End If
    
    goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Sean McAvoy
'Date: 12-18-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset
Const sPROC_NAME = "clsExProsOccl.LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    If mvarDentalEncounterID = 0 Then
        clsISection_LoadSection = False
        Exit Function
    End If
    'Load the values
    sSQL = "SELECT * FROM TEMPL_EXAM_PROS_OCCLUSAL Where " + _
        "DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        mstrCentricRelation = ""
        mstrLeavesAtContact = ""
        mstrPointOfContact = ""
        mstrMaterialUsed = ""
        mstrCarrierType = ""
        msinA_P_Slide = -1
        msinLateralSlide = -1
        msinRtLateralGuidance = -1
        msinLtLateralGuidance = -1
        mstrProtrusiveGuidance = ""
        mstrRtAnteriorCoupling = ""
        mstrLtAnteriorCoupling = ""
        mstrCRreproducibility = ""
   
    Else
        dfDirtyFlag = dfNoAction
        
        mstrCentricRelation = rs("CENTRICRELATION") & ""
        mstrLeavesAtContact = rs("LEAVESATCONTACT") & ""
        mstrPointOfContact = rs("POINTOFCONTACT") & ""
        mstrMaterialUsed = rs("MATERIALUSED") & ""
        mstrCarrierType = rs("CARRIERTYPE") & ""
        
        '***** Need to know if these number values are Null or not
        If IsNull(rs("A_P_SLIDE")) = True Then
            msinA_P_Slide = -1
        Else
            msinA_P_Slide = Val(rs("A_P_SLIDE") & "")
        End If
        If IsNull(rs("LATERALSLIDE")) = True Then
            msinLateralSlide = -1
        Else
            msinLateralSlide = Val(rs("LATERALSLIDE") & "")
        End If
        If IsNull(rs("RTLATGUIDANCE")) = True Then
            msinRtLateralGuidance = -1
        Else
            msinRtLateralGuidance = Val(rs("RTLATGUIDANCE") & "")
        End If
        If IsNull(rs("LTLATGUIDANCE")) = True Then
            msinLtLateralGuidance = -1
        Else
            msinLtLateralGuidance = Val(rs("LTLATGUIDANCE") & "")
        End If
        '***********
        mstrProtrusiveGuidance = rs("PROTGUIDANCE") & ""
        mstrRtAnteriorCoupling = rs("RTANTERIORCUPLING") & ""
        mstrLtAnteriorCoupling = rs("LTANTERIORCUPLING") & ""
        mstrCRreproducibility = rs("CRREPRODUCIBILITY") & ""
    End If
    
    rs.Close
    
    
    Set rs = Nothing
    'SM. 01/29/02
    'The format of the data for "Material Used" has changed.
    'The following line is used to check the presense of that old data.
    'If old data found, clear it. This line can be deleted once all old data removed.
    If Len(mstrMaterialUsed) > 0 Then
        If InStr(1, mstrMaterialUsed, "^", vbTextCompare) = 0 Then
            mstrMaterialUsed = ""
        End If
    End If
    
    clsISection_LoadSection = True
    
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
    
End Function

Private Function UpdateDirtyFlag()

    If dfDirtyFlag = dfNoAction Then
        dfDirtyFlag = dfUpdate
    End If
    
End Function

Private Sub Class_Initialize()

    dfDirtyFlag = dfNoAction
    
    mstrCentricRelation = ""
    mstrLeavesAtContact = ""
    mstrPointOfContact = ""
    mstrMaterialUsed = ""
    mstrCarrierType = ""
    msinA_P_Slide = -1
    msinLateralSlide = -1
    msinRtLateralGuidance = -1
    msinLtLateralGuidance = -1
    mstrProtrusiveGuidance = ""
    mstrRtAnteriorCoupling = ""
    mstrLtAnteriorCoupling = ""
    mstrCRreproducibility = ""
    
    colMaterialUsed.Add "Polyvinylsiloxane"
    colMaterialUsed.Add "Polyether"
    colMaterialUsed.Add "Zinc Oxide"
    colMaterialUsed.Add "Plaster"
    colMaterialUsed.Add "Modeling Plastic"
    colMaterialUsed.Add "Wax"
    colMaterialUsed.Add "Acrylic Resin"

    
    
End Sub






