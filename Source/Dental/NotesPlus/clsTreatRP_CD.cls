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
Attribute VB_Name = "clsTreatRP_CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Dental Templates Treatment - Removable Pros - Complete Denture tab

Implements clsISection

Private mvarPatientID As Long
Private mvarDentalEncounterID As Long

Private dfDirtyFlag As DirtyFlagEnum
'local variable(s) to hold property value(s)
Private mstrTreatment_Man As String
Private mstrTreatment_Max As String
Private mstrOcclusal_Scheme As String
Private mstrMaterial_MaxBase As String
Private mstrMaterial_ManBase As String
Private mstrShade_MaxBase As String
Private mstrShade_ManBase As String
Private mintWaxTryin As Integer
Private mintAccepts_Esthetics As Integer
Private mintDeliver As Integer
Private mintPost_Insertion As Integer
Private mintClinical_Remount As Integer
Private mintPostOp_Adjustment As Integer
Private mstrComments As String
Private mstrSupporting_Procedures As String
Private mstrTreatment_Man_Comments As String
Private mstrTreatment_Max_Comments As String



Public Property Let Supporting_Procedures(ByVal vData As String)
    mstrSupporting_Procedures = vData
    UpdateDirtyFlag
End Property

Public Property Get Supporting_Procedures() As String

    If mstrSupporting_Procedures = "" Then
        mstrSupporting_Procedures = gcSupportingProcs
    End If
    
    Supporting_Procedures = mstrSupporting_Procedures

End Property

Public Property Let Comments(ByVal vData As String)
    mstrComments = vData
    UpdateDirtyFlag
End Property

Public Property Get Comments() As String
    Comments = mstrComments
End Property

Public Property Let PostOp_Adjustment(ByVal vData As Integer)
    mintPostOp_Adjustment = vData
    UpdateDirtyFlag
End Property

Public Property Get PostOp_Adjustment() As Integer
    PostOp_Adjustment = mintPostOp_Adjustment
End Property

Public Property Let Clinical_Remount(ByVal vData As Integer)
    mintClinical_Remount = vData
    UpdateDirtyFlag
End Property

Public Property Get Clinical_Remount() As Integer
    Clinical_Remount = mintClinical_Remount
End Property

Public Property Let Post_Insertion(ByVal vData As Integer)
    mintPost_Insertion = vData
    UpdateDirtyFlag
End Property

Public Property Get Post_Insertion() As Integer
    Post_Insertion = mintPost_Insertion
End Property

Public Property Let Deliver(ByVal vData As Integer)
    mintDeliver = vData
    UpdateDirtyFlag
End Property

Public Property Get Deliver() As Integer
    Deliver = mintDeliver
End Property

Public Property Let Accepts_Esthetics(ByVal vData As Integer)
    mintAccepts_Esthetics = vData
    UpdateDirtyFlag
End Property

Public Property Get Accepts_Esthetics() As Integer
    Accepts_Esthetics = mintAccepts_Esthetics
End Property

Public Property Let WaxTryin(ByVal vData As Integer)
    mintWaxTryin = vData
    UpdateDirtyFlag
End Property

Public Property Get WaxTryin() As Integer
    WaxTryin = mintWaxTryin
End Property

Public Property Let Shade_ManBase(ByVal vData As String)
    mstrShade_ManBase = vData
    UpdateDirtyFlag
End Property

Public Property Get Shade_ManBase() As String
    Shade_ManBase = mstrShade_ManBase
End Property

Public Property Let Shade_MaxBase(ByVal vData As String)
    mstrShade_MaxBase = vData
    UpdateDirtyFlag
End Property

Public Property Get Shade_MaxBase() As String
    Shade_MaxBase = mstrShade_MaxBase
End Property

Public Property Let Material_ManBase(ByVal vData As String)
    mstrMaterial_ManBase = vData
    UpdateDirtyFlag
End Property

Public Property Get Material_ManBase() As String
    Material_ManBase = mstrMaterial_ManBase
End Property

Public Property Let Material_MaxBase(ByVal vData As String)
    mstrMaterial_MaxBase = vData
    UpdateDirtyFlag
End Property

Public Property Get Material_MaxBase() As String
    Material_MaxBase = mstrMaterial_MaxBase
End Property

Public Property Let Occlusal_Scheme(ByVal vData As String)
    mstrOcclusal_Scheme = vData
    UpdateDirtyFlag
End Property

Public Property Get Occlusal_Scheme() As String
    Occlusal_Scheme = mstrOcclusal_Scheme
End Property


Public Property Let Treatment_Max_Comments(ByVal vData As String)
    mstrTreatment_Max_Comments = vData
    UpdateDirtyFlag
End Property

Public Property Get Treatment_Max_Comments() As String
    Treatment_Max_Comments = mstrTreatment_Max_Comments
    UpdateDirtyFlag
End Property



Public Property Let Treatment_Max(ByVal vData As String)
    mstrTreatment_Max = vData
    UpdateDirtyFlag
End Property

Public Property Get Treatment_Max() As String
    Treatment_Max = mstrTreatment_Max
    UpdateDirtyFlag
End Property

Public Property Let Treatment_Man_Comments(ByVal vData As String)
    mstrTreatment_Man_Comments = vData
    UpdateDirtyFlag
End Property

Public Property Get Treatment_Man_Comments() As String
    Treatment_Man_Comments = mstrTreatment_Man_Comments
End Property

Public Property Let Treatment_Man(ByVal vData As String)
    mstrTreatment_Man = vData
    UpdateDirtyFlag
End Property

Public Property Get Treatment_Man() As String
    Treatment_Man = mstrTreatment_Man
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
On Error GoTo ErrHnd
Const sPROC_NAME As String = "clsTreatRP_CD.clsISection_GetRTFNote"
    
    If mstrTreatment_Man = "" And _
        mstrTreatment_Max = "" And _
        mstrOcclusal_Scheme = "" And _
        mstrMaterial_MaxBase = "" And _
        mstrMaterial_ManBase = "" And _
        mstrShade_MaxBase = "" And _
        mstrShade_ManBase = "" And _
        mintWaxTryin = 0 And _
        mintAccepts_Esthetics = 0 And _
        mintDeliver = 0 And _
        mintPost_Insertion = 0 And _
        mintClinical_Remount = 0 And _
        mintPostOp_Adjustment = 0 And _
        mstrComments = "" And _
        mstrTreatment_Man_Comments = "" And _
        mstrTreatment_Max_Comments = "" And _
        mstrSupporting_Procedures = "" Then
        
        clsISection_GetRTFNote = ""
        Exit Function
    End If
    
    If mstrTreatment_Man <> "" Or mstrTreatment_Max <> "" Or _
            mstrTreatment_Max_Comments <> "" Or mstrTreatment_Man_Comments <> "" Then
        sNote = sNote + crtfStartControlName & "Treatment Info: "
        If mstrTreatment_Max <> "" Then
            sTxt = "Max> " + Replace(mstrTreatment_Max, ":Y", ":Y; ")
            sTxt = Replace(sTxt, ":N", ":N; ")
            sNote = sNote & sTxt
            sTxt = ""
        End If
        If mstrTreatment_Max_Comments <> "" Then
            sNote = sNote & "Comments: " + mstrTreatment_Max_Comments + "; "
        End If
        If mstrTreatment_Man <> "" Then
            sTxt = "Man> " + Replace(mstrTreatment_Man, ":Y", ":Y; ")
            sTxt = Replace(sTxt, ":N", ":N; ")
            sNote = sNote + sTxt
            sTxt = ""
        End If
        If mstrTreatment_Man_Comments <> "" Then
            sNote = sNote + "Comments: " + mstrTreatment_Man_Comments + "; "
        End If
        
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote + crtfEndControlName
    End If
    
    If mstrOcclusal_Scheme <> "" Or mstrMaterial_MaxBase <> "" Or mstrMaterial_ManBase <> "" Or _
        mstrShade_MaxBase <> "" Or mstrShade_ManBase <> "" Or mintWaxTryin = -1 Or mintAccepts_Esthetics = -1 Then
        sNote = sNote + crtfStartControlName & "Occlusion, Materials Selection & Try-in: "
        If mstrOcclusal_Scheme <> "" Then _
            sNote = sNote + "Occlusal Scheme: " + mstrOcclusal_Scheme + "; "
        If mstrMaterial_MaxBase <> "" Then _
            sNote = sNote + "Max Base Material: " + mstrMaterial_MaxBase + "; "
        If mstrShade_MaxBase <> "" Then _
            sNote = sNote + "Max Base Shade: " + mstrShade_MaxBase + "; "
        If mstrMaterial_ManBase <> "" Then _
            sNote = sNote + "Man Base Material: " + mstrMaterial_ManBase + "; "
        If mstrShade_ManBase <> "" Then _
            sNote = sNote + "Man Base Shade: " + mstrShade_ManBase + "; "
        
        If mintWaxTryin = -1 Then _
            sNote = sNote + "Wax Try-in" + "; "
        If mintAccepts_Esthetics = -1 Then _
            sNote = sNote + "Patient accepts esthetics" + "; "
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote + crtfEndControlName
    End If
    
    If mintDeliver = -1 Or mintPost_Insertion = -1 Or mintClinical_Remount = -1 Or _
     mintPostOp_Adjustment = -1 Then
        sNote = sNote + crtfStartControlName & "Complete Denture Delivery: "
        If mintDeliver = -1 Then _
            sNote = sNote & "Deliver; "
        If mintPost_Insertion = -1 Then _
            sNote = sNote + "Post-insertion Instructions; "
        If mintClinical_Remount = -1 Then _
            sNote = sNote + "Clinical Remount; "
        If mintPostOp_Adjustment = -1 Then _
            sNote = sNote + "Post-op Adjustment; "
        'Remove the last "; "
        sNote = Left(sNote, Len(sNote) - Len("; "))
        sNote = sNote + crtfEndControlName
    End If

    sTxt = SupportingProcsDisplay
    If sTxt <> "" Or mstrComments <> "" Then
        If sTxt <> "" Then
            sNote = sNote & crtfStartControlName & "Supporting Procedures: " & sTxt + crtfEndControlName
        End If
        
        If mstrComments <> "" Then _
            sNote = sNote + crtfStartControlName & "Complete Denture Comments: " + mstrComments + crtfEndControlName
    End If
    clsISection_GetRTFNote = sNote
    Exit Function
ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function

'*******************************************************************
'Function: GetNote
'Author: Sean McAvoy
'Date: 01/08/02
'Returns: A string that is displayed in the narrative summary for this encounter
'Purpose: To compare the values from this encounter and the previous encounter and
'return a text note.
'*******************************************************************
Private Function clsISection_GetNote() As String

Dim sNote As String
Dim sTxt As String
On Error GoTo ErrHnd
Const sPROC_NAME As String = "clsTreatRP_CD.clsISection_GetNote"
    
    If mstrTreatment_Man = "" And _
        mstrTreatment_Max = "" And _
        mstrOcclusal_Scheme = "" And _
        mstrMaterial_MaxBase = "" And _
        mstrMaterial_ManBase = "" And _
        mstrShade_MaxBase = "" And _
        mstrShade_ManBase = "" And _
        mintWaxTryin = 0 And _
        mintAccepts_Esthetics = 0 And _
        mintDeliver = 0 And _
        mintPost_Insertion = 0 And _
        mintClinical_Remount = 0 And _
        mintPostOp_Adjustment = 0 And _
        mstrComments = "" And _
        mstrSupporting_Procedures = "" Then
        
        clsISection_GetNote = ""
        Exit Function
    End If

    sNote = vbCrLf + "Treatment - Removable Pros - Complete Dentures" + vbCrLf
    If mstrTreatment_Man <> "" Then
        sTxt = Replace(mstrTreatment_Man, ":Y", ":Y   ")
        sTxt = Replace(sTxt, ":N", ":N   ")
        sNote = sNote + vbTab + "Mandibular Treatment: " + sTxt + vbCrLf
        sTxt = ""
    End If
    If mstrTreatment_Max <> "" Then
        sTxt = Replace(mstrTreatment_Max, ":Y", ":Y   ")
        sTxt = Replace(sTxt, ":N", ":N   ")
        sNote = sNote + vbTab + "Maxillary Treatment: " + sTxt + vbCrLf
        sTxt = ""
    End If
    If mstrOcclusal_Scheme <> "" Then _
        sNote = sNote + vbTab + "Occlusal Scheme: " + mstrOcclusal_Scheme + vbCrLf
    
    If mintWaxTryin = -1 Then _
        sNote = sNote + vbTab + "WaxTry-in" + vbCrLf
    If mintAccepts_Esthetics = -1 Then _
        sNote = sNote + vbTab + "Accepts Esthetics" + vbCrLf
    If mintDeliver = -1 Then _
        sNote = sNote + vbTab + "Deliver" + vbCrLf
    If mintPost_Insertion = -1 Then _
        sNote = sNote + vbTab + "Post Insertion Instructions" + vbCrLf
    If mintClinical_Remount = -1 Then _
        sNote = sNote + vbTab + "Clinical Remount" + vbCrLf
    If mintPostOp_Adjustment = -1 Then _
        sNote = sNote + vbTab + "Post-op Adjustment" + vbCrLf

    sTxt = SupportingProcsDisplay
    If sTxt <> "" Then
        sNote = sNote & vbTab & "Supporting Procedures: " & sTxt & vbCrLf
    End If
    If mstrMaterial_MaxBase <> "" Then _
        sNote = sNote + vbTab + "Maxillary Base Material Info: " + mstrMaterial_MaxBase + vbCrLf
    If mstrMaterial_ManBase <> "" Then _
        sNote = sNote + vbTab + "Mandibular Base Material Info: " + mstrMaterial_ManBase + vbCrLf
    If mstrShade_MaxBase <> "" Then _
        sNote = sNote + vbTab + "Maxillary Base Shade Info: " + mstrShade_MaxBase + vbCrLf
    If mstrShade_ManBase <> "" Then _
        sNote = sNote + vbTab + "Mandibular Base Shade Info: " + mstrShade_ManBase + vbCrLf
    
    If mstrComments <> "" Then _
        sNote = sNote + vbTab + "Complete Denture Comments: " + mstrComments + vbCrLf
        
    clsISection_GetNote = sNote
    Exit Function
ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical
End Function
'*******************************************************************
'Function: SupportingProcsDisplay
'Author: Sean McAvoy
'Date: 1/2/02
'Returns: String
'Purpose: To get the display string for supporting procedures
'*******************************************************************
Private Function SupportingProcsDisplay() As String

Dim oItems As New CMultiListItems
'Dim oItem As CMultiListItem

    oItems.ListType = Standard
    oItems.LoadFromValueString mstrSupporting_Procedures
    SupportingProcsDisplay = oItems.Value_String_Display
Set oItems = Nothing
'Set oItem = Nothing
End Function

'*******************************************************************
'Function: SaveSection
'Author: Sean McAvoy
'Date: 01/08/02
'Returns: True/False
'Purpose: Saves the data to the table
'*******************************************************************
Private Function clsISection_SaveSection() As Boolean

Dim sSQL As String
Dim mstrMaterial_MaxBase_TMP As String
Dim mstrMaterial_ManBase_TMP As String
Dim mstrShade_MaxBase_TMP As String
Dim mstrShade_ManBase_TMP As String
Dim mstrComments_TMP As String
Dim mstrTreatment_Man_Comments_TMP As String
Dim mstrTreatment_Max_Comments_TMP As String

Const sPROC_NAME As String = "clsTreatRP_CD.clsISection_SaveSection"

On Error GoTo ErrHnd

    If dfDirtyFlag = dfNoAction Then
        clsISection_SaveSection = True
        Exit Function
    End If
        
    mstrMaterial_MaxBase_TMP = RQ(mstrMaterial_MaxBase)
    mstrMaterial_ManBase_TMP = RQ(mstrMaterial_ManBase)
    mstrShade_MaxBase_TMP = RQ(mstrShade_MaxBase)
    mstrShade_ManBase_TMP = RQ(mstrShade_ManBase)
    mstrComments_TMP = RQ(mstrComments)
    mstrTreatment_Man_Comments_TMP = RQ(mstrTreatment_Man_Comments)
    mstrTreatment_Max_Comments_TMP = RQ(mstrTreatment_Max_Comments)
        
    If dfDirtyFlag = dfAdd Then
        'The record does not exist, so create an insert statement
        sSQL = "INSERT INTO TEMPL_TREAT_RP_CD (" _
            & " DENTAL_ENCOUNTER_ID, TREATMENT_MAN, TREATMENT_MAX, OCCLUSAL_SCHEME, MATERIAL_MAXBASE," _
            & " MATERIAL_MANBASE, SHADE_MAXBASE, SHADE_MANBASE, WAXTRYIN, ACCEPTS_ESTHETICS," _
            & " DELIVER, POST_INSERTION, CLINICAL_REMOUNT, POSTOP_ADJUSTMENT, COMMENTS," _
            & " SUPPORTING_PROCEDURES, TREATMENT_MAN_COMMENTS, TREATMENT_MAX_COMMENTS) " _
            & " VALUES" _
            & " (" & CStr(mvarDentalEncounterID) & ", " _
            & " '" & mstrTreatment_Man & "', " _
            & " '" & mstrTreatment_Max & "', " _
            & " '" & mstrOcclusal_Scheme & "', " _
            & " '" & mstrMaterial_MaxBase_TMP & "', " _
            & " '" & mstrMaterial_ManBase_TMP & "', " _
            & " '" & mstrShade_MaxBase_TMP & "', " _
            & " '" & mstrShade_ManBase_TMP & "', " _
            & CStr(mintWaxTryin) & ", " _
            & CStr(mintAccepts_Esthetics) & ", " _
            & CStr(mintDeliver) & ", " _
            & CStr(mintPost_Insertion) & ", " _
            & CStr(mintClinical_Remount) & ", " _
            & CStr(mintPostOp_Adjustment) & ", " _
            & " '" & mstrComments_TMP & "', " _
            & " '" & mstrSupporting_Procedures & "', " _
            & " '" & mstrTreatment_Man_Comments_TMP & "', " _
            & " '" & mstrTreatment_Max_Comments_TMP & "') " _

    ElseIf dfDirtyFlag = dfUpdate Then
        'The record exists, so create an update statement
        sSQL = "UPDATE TEMPL_TREAT_RP_CD SET " + _
                "TREATMENT_MAN = '" + mstrTreatment_Man + "', " + _
                "TREATMENT_MAX = '" + mstrTreatment_Max + "', " + _
                "OCCLUSAL_SCHEME = '" + mstrOcclusal_Scheme + "', " + _
                "MATERIAL_MAXBASE = '" + mstrMaterial_MaxBase_TMP + "', " + _
                "MATERIAL_MANBASE = '" + mstrMaterial_ManBase_TMP + "', " + _
                "SHADE_MAXBASE = '" + mstrShade_MaxBase_TMP + "', " + _
                "SHADE_MANBASE = '" + mstrShade_ManBase_TMP + "', " + _
                "WAXTRYIN = " + CStr(mintWaxTryin) + ", " + _
                "ACCEPTS_ESTHETICS = " + CStr(mintAccepts_Esthetics) + ", " + _
                "DELIVER = " + CStr(mintDeliver) + ", " + _
                "POST_INSERTION = " + CStr(mintPost_Insertion) + ", " + _
                "CLINICAL_REMOUNT = " + CStr(mintClinical_Remount) + ", " + _
                "POSTOP_ADJUSTMENT = " + CStr(mintPostOp_Adjustment) + ", " + _
                "COMMENTS = '" + mstrComments_TMP + "', " + _
                "SUPPORTING_PROCEDURES = '" + mstrSupporting_Procedures + "', " + _
                "TREATMENT_MAN_COMMENTS = '" + mstrTreatment_Man_Comments_TMP + "', " + _
                "TREATMENT_MAX_COMMENTS = '" + mstrTreatment_Max_Comments_TMP + "' " + _
                "Where (DENTAL_ENCOUNTER_ID = " & CStr(mvarDentalEncounterID) & ")"
        
    ElseIf dfDirtyFlag = dfdelete Then
        'Mark the record as deleted
        sSQL = ""
    End If
    
    If sSQL <> "" Then goconn.ExecuteSQL sSQL
    
    dfDirtyFlag = dfNoAction
    
    clsISection_SaveSection = True

Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function


'*******************************************************************
'Function: LoadSection
'Author: Sean McAvoy
'Date: 10-22-01
'Returns: True/False
'Purpose: This loads the data from the db
'*******************************************************************
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean

Dim sSQL As String, rs As Recordset, rsPrev As Recordset
Const sPROC_NAME = "clsTreatRP_CD.clsISection_LoadSection"

On Error GoTo ErrHnd

    mvarPatientID = PatientID
    mvarDentalEncounterID = DentalEncounterID
    If mvarDentalEncounterID = 0 Then
        clsISection_LoadSection = False
        Exit Function
    End If
    
    'Load the values
    'Because of the issues with Tuxedo (4k limit), three sql calls will
    'need to be made to retrieve the several 2k note fields.
    
    sSQL = "SELECT DENTAL_ENCOUNTER_ID, TREATMENT_MAN, TREATMENT_MAX, OCCLUSAL_SCHEME, MATERIAL_MAXBASE, " + _
           "MATERIAL_MANBASE, WAXTRYIN, ACCEPTS_ESTHETICS, DELIVER, POST_INSERTION, CLINICAL_REMOUNT, " + _
           "POSTOP_ADJUSTMENT, COMMENTS, SUPPORTING_PROCEDURES FROM TEMPL_TREAT_RP_CD " + _
           "WHERE  DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)

    Set rs = goconn.OpenRecordset(sSQL)
    
    If rs.EOF Then
        dfDirtyFlag = dfAdd
        mstrTreatment_Man = ""
        mstrTreatment_Max = ""
        mstrOcclusal_Scheme = ""
        mstrMaterial_MaxBase = ""
        mstrMaterial_ManBase = ""
        mstrShade_MaxBase = ""
        mstrShade_ManBase = ""
        mintWaxTryin = 0
        mintAccepts_Esthetics = 0
        mintDeliver = 0
        mintPost_Insertion = 0
        mintClinical_Remount = 0
        mintPostOp_Adjustment = 0
        mstrComments = ""
        mstrSupporting_Procedures = ""
        mstrTreatment_Man_Comments = ""
        mstrTreatment_Max_Comments = ""

    Else
        dfDirtyFlag = dfNoAction
        
        mstrTreatment_Man = rs("TREATMENT_MAN") & ""
        mstrTreatment_Max = rs("TREATMENT_MAX") & ""
        mstrOcclusal_Scheme = rs("OCCLUSAL_SCHEME") & ""
        mstrMaterial_MaxBase = rs("MATERIAL_MAXBASE") & ""
        mstrMaterial_ManBase = rs("MATERIAL_MANBASE") & ""
        mintWaxTryin = Val(rs("WAXTRYIN") & "")
        mintAccepts_Esthetics = Val(rs("ACCEPTS_ESTHETICS") & "")
        mintDeliver = Val(rs("DELIVER") & "")
        mintPost_Insertion = Val(rs("POST_INSERTION") & "")
        mintClinical_Remount = Val(rs("CLINICAL_REMOUNT") & "")
        mintPostOp_Adjustment = Val(rs("POSTOP_ADJUSTMENT") & "")
        rs.Close
        
        
        sSQL = "SELECT SHADE_MAXBASE, SHADE_MANBASE  FROM TEMPL_TREAT_RP_CD " + _
           "WHERE  DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
        Set rs = goconn.OpenRecordset(sSQL)
        If rs.EOF Then
            mstrShade_MaxBase = ""
            mstrShade_ManBase = ""
        Else
            mstrShade_MaxBase = rs("SHADE_MAXBASE") & ""
            mstrShade_ManBase = rs("SHADE_MANBASE") & ""
        End If
        rs.Close
        
        sSQL = "SELECT COMMENTS, SUPPORTING_PROCEDURES FROM TEMPL_TREAT_RP_CD " + _
           "WHERE  DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
        Set rs = goconn.OpenRecordset(sSQL)
        
        If rs.EOF Then
            mstrComments = ""
            mstrSupporting_Procedures = ""
        Else
            mstrComments = rs("COMMENTS") & ""
            mstrSupporting_Procedures = rs("SUPPORTING_PROCEDURES") & ""
        End If
        rs.Close
        
        sSQL = "SELECT TREATMENT_MAN_COMMENTS, TREATMENT_MAX_COMMENTS FROM TEMPL_TREAT_RP_CD " + _
           "WHERE  DENTAL_ENCOUNTER_ID = " + CStr(mvarDentalEncounterID)
        Set rs = goconn.OpenRecordset(sSQL)
        
        If rs.EOF Then
            mstrTreatment_Man_Comments = ""
            mstrTreatment_Max_Comments = ""
        Else
            mstrTreatment_Man_Comments = rs("TREATMENT_MAN_COMMENTS") & ""
            mstrTreatment_Max_Comments = rs("TREATMENT_MAX_COMMENTS") & ""
        End If
        rs.Close
        
    End If
    
    Set rs = Nothing
    
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
    mstrTreatment_Man = ""
    mstrTreatment_Max = ""
    mstrOcclusal_Scheme = ""
    mstrMaterial_MaxBase = ""
    mstrMaterial_ManBase = ""
    mstrShade_MaxBase = ""
    mstrShade_ManBase = ""
    mintWaxTryin = 0
    mintAccepts_Esthetics = 0
    mintDeliver = 0
    mintPost_Insertion = 0
    mintClinical_Remount = 0
    mintPostOp_Adjustment = 0
    mstrComments = ""
    mstrSupporting_Procedures = ""
    mstrTreatment_Man_Comments = ""
    mstrTreatment_Max_Comments = ""
    
    
End Sub






