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
Attribute VB_Name = "colProsImpress"
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
Private mcolImpressions As Collection

Private Const mcDiagImpressionsWorkload As Long = 117
Private Const const_ModuleName = "colProsImpressions" '*** Insert Module Name Here

Private Sub Class_Initialize()
    Set mcolImpressions = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolImpressions = Nothing
End Sub

Public Property Get Col_Location() As Long
    Col_Location = 0
End Property
Public Property Get Col_Purpose() As Long
    Col_Purpose = 1
End Property
Public Property Get Col_NumMade() As Long
    Col_NumMade = 2
End Property
Public Property Get Col_Material() As Long
    Col_Material = 3
End Property

'*****************************************************************************
'Function: BuildLocationList
'Author:KDunne
'Date:1/8/2002
'Purpose: Builds a list of Locations for the combo list
'*****************************************************************************
Public Function BuildLocationList() As String
    Dim strReturn As String
    Dim objImpress As clsProsImpress
    Dim X As Integer
    
    strReturn = Empty
    
    On Error GoTo ErrHandler
    
    Set objImpress = New clsProsImpress
    
    For X = 0 To 2
        strReturn = strReturn & objImpress.GetProsImpressLocationEnumDesc(X) & " |"
    Next
    
FuncExit:
    BuildLocationList = strReturn
    Set objImpress = Nothing
    Exit Function
ErrHandler:
    strReturn = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".BuildLocationList()", "", vbCritical
    GoTo FuncExit
End Function

'*****************************************************************************
'Function: BuildPurposeList
'Author:KDunne
'Date:1/8/2002
'Purpose: Builds a list of Purposes for the combo list
'*****************************************************************************
Public Function BuildPurposeList() As String
    Dim strReturn As String
    Dim objImpress As clsProsImpress
    Dim X As Integer
    
    strReturn = Empty
    
    On Error GoTo ErrHandler
    Set objImpress = New clsProsImpress
    
    For X = 0 To 4
        strReturn = strReturn & objImpress.GetProsImpressPurposeEnumDesc(X) & " |"
    Next
    
FuncExit:
    BuildPurposeList = strReturn
    Set objImpress = Nothing
    Exit Function
ErrHandler:
    strReturn = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".BuildPurposeList()", "", vbCritical
    GoTo FuncExit
End Function

'*****************************************************************************
'Function: BuildMaterialList
'Author:KDunne
'Date:1/8/2002
'Purpose: Builds a list of Materials for the combo list
'*****************************************************************************
Public Function BuildMaterialList() As String
    Dim strReturn As String
    Dim objImpress As clsProsImpress
    Dim X As Integer
    
    strReturn = Empty
    
    On Error GoTo ErrHandler
    Set objImpress = New clsProsImpress
    
    For X = 0 To 5
        strReturn = strReturn & objImpress.GetProsImpressMaterialEnumDesc(X) & " |"
    Next
    
FuncExit:
    BuildMaterialList = strReturn
    Set objImpress = Nothing
    Exit Function
ErrHandler:
    strReturn = Empty
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".BuildMaterialList()", "", vbCritical
    GoTo FuncExit
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mcolImpressions.[_NewEnum]
End Property
Public Property Get Count() As Long
    Count = mcolImpressions.Count
End Property

Public Property Get Impression(ByVal intIndex As Long) As clsProsImpress
    Dim objImpress As clsProsImpress

    If intIndex > mcolImpressions.Count Then
        Set objImpress = New clsProsImpress
        objImpress.Dental_EncounterId = mlngDentalEncounterId
        mcolImpressions.Add objImpress
    End If
    
    Set Impression = mcolImpressions(intIndex)
    
    Set objImpress = Nothing
End Property

Private Property Get clsISection_DentalEncounterID() As Long
    clsISection_DentalEncounterID = mlngDentalEncounterId
End Property

Private Function clsISection_GetNote() As String
    Dim strNote As String
    Dim objImpress As clsProsImpress
    
    strNote = Empty
    
    For Each objImpress In mcolImpressions
        strNote = strNote & objImpress.GetRTFNote() & vbCrLf
    Next
    
    If strNote <> "" Then
        strNote = crtfStartControlName & "Prosthodontic Impressions" & _
            strNote & crtfEndControlName
    End If
    clsISection_GetNote = strNote
    Set objImpress = Nothing
End Function
Private Function clsISection_GetRTFNote() As String
    Dim strNote As String
    Dim objImpress As clsProsImpress
    Dim sTxt As String
    
    strNote = Empty
    
    For Each objImpress In mcolImpressions
        sTxt = objImpress.GetRTFNote()
        If sTxt <> "" Then
            strNote = strNote & sTxt & "  "
            sTxt = ""
        End If
'        strNote = strNote & objImpress.GetNote() & vbCrLf
    Next
    
'    If strNote <> "" Then
'        strNote = crtfStartControlName & "Prosthodontic Impressions: " & _
'            strNote & crtfEndControlName
'    End If
    clsISection_GetRTFNote = strNote
    Set objImpress = Nothing
End Function
Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
    Dim blnReturn As Boolean
    Dim objRS As ADODB.Recordset
    Dim objImpress As clsProsImpress
    
    On Error GoTo ErrHandler
    
    mlngPatUnitId = PatientID
    mlngDentalEncounterId = DentalEncounterID
    
    Set mcolImpressions = New Collection
    
    Set objRS = DB_Load_Impressions(ByVal mlngDentalEncounterId)
    
    With objRS
        If .EOF Then
            mlngDirtyFlag = dfAdd
            'reset property values here
        
        Else
            mlngDirtyFlag = dfUpdate
            Do While Not .EOF
                Set objImpress = New clsProsImpress
                
                objImpress.Dental_EncounterId = mlngDentalEncounterId
                objImpress.ImpressionId = .Fields("IMPRESSION_ID")
                objImpress.Location = .Fields("LOCATION")
                objImpress.Number_Made = .Fields("NUM_MADE")
                objImpress.Purpose = .Fields("PURPOSE")
                objImpress.Material = .Fields("Material")
                objImpress.WorkloadID = .Fields("Workload_ID")
                objImpress.Clear_Dirty_Flag
                
                mcolImpressions.Add objImpress
            
                .MoveNext
            Loop
        End If
    End With
    
FuncExit:
    clsISection_LoadSection = blnReturn
    Set objRS = Nothing
    Exit Function
ErrHandler:
   blnReturn = False
   goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_LoadSection()", "", vbCritical
   GoTo FuncExit
End Function

Private Property Get clsISection_PatientID() As Long
    clsISection_PatientID = mlngPatUnitId
End Property

Private Function clsISection_SaveSection() As Boolean
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim objImpression As clsProsImpress
    
    On Error GoTo ErrHandler
    
    DetermineWorkload

    For Each objImpression In mcolImpressions
        objImpression.SaveSection
    Next
    
    blnReturn = True
FuncExit:
    clsISection_SaveSection = blnReturn
    Exit Function
    
ErrHandler:
    blnReturn = False
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".clsISection_SaveSection()", "", vbCritical
    GoTo FuncExit
End Function

Private Sub DetermineWorkload()

    '6/23/03 - 36294
    'Assign workload once per arch, twice per encounter (for two arches)
    
    Dim oImpress As clsProsImpress
    Dim oWorkloadImpress As clsProsImpress
    Dim bAppliedWorkload As Boolean
    Dim bApplyWorkload As Boolean
    Dim lWorkloadID As Long
    Dim iManMultiplier As Integer
    Dim iMaxMultiplier As Integer

    iManMultiplier = 0
    iMaxMultiplier = 0
    
    For Each oImpress In mcolImpressions
        
        If oImpress.WorkloadID <> 0 Then
            'Workload has been applied to this object, so remember it (in case we need to delete it) and remove it from this object
            bAppliedWorkload = True
            lWorkloadID = oImpress.WorkloadID
            oImpress.WorkloadID = 0
        End If
        
        If oImpress.Purpose = piPur_Diagnostic Then
            If oImpress.Location = piLoc_Mandibular Then
                iManMultiplier = 1
            ElseIf oImpress.Location = piLoc_Maxillary Then
                iMaxMultiplier = 1
            End If
            
            'Workload should be applied to this object
            bApplyWorkload = True
            Set oWorkloadImpress = oImpress

        End If
    Next
    
    If bApplyWorkload = True And bAppliedWorkload = False Then
        'Add the workload
        UpdateWorkload lWorkloadID, dfAdd, mlngDentalEncounterId, mcDiagImpressionsWorkload, iManMultiplier + iMaxMultiplier
        oWorkloadImpress.WorkloadID = lWorkloadID
        
    ElseIf bApplyWorkload = True And bAppliedWorkload = True Then
        'Update the workload
        UpdateWorkload lWorkloadID, dfUpdate, mlngDentalEncounterId, mcDiagImpressionsWorkload, iManMultiplier + iMaxMultiplier
        oWorkloadImpress.WorkloadID = lWorkloadID
    
    ElseIf bApplyWorkload = False And bAppliedWorkload = True Then
        'Delete the workload
        UpdateWorkload lWorkloadID, dfdelete, mlngDentalEncounterId, mcDiagImpressionsWorkload, iManMultiplier + iMaxMultiplier
    End If
    
    Set oWorkloadImpress = Nothing
    
End Sub


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

'*********************************************************************************
'Function: DB_Load_Impressions
'Author: KDunne
'Date: 1/7/2002
'Purpose: Loads the impressions from the table TEMP_GENPROS_IMPRESSIONS
'*********************************************************************************
Private Function DB_Load_Impressions(ByVal lngDentalEncounterId As Long) As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrHandler
    
    strSql = "Select IMPRESSION_ID, LOCATION, NUM_MADE , Material, PURPOSE, Workload_ID " & _
             " FROM TEMPL_GENPROS_IMPRESS WHERE DENTAL_ENCOUNTER_ID = " & lngDentalEncounterId

    Set DB_Load_Impressions = goconn.OpenRecordset(strSql)

    Exit Function
ErrHandler:
    goCWShared.ShowVBError Err.Number, Err.Description, const_ModuleName & ".DB_Load_Impressions()", "", vbCritical
    Set DB_Load_Impressions = Nothing
End Function
