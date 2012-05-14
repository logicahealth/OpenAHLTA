VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private moDAS As CHCSII_DAS
'Private moDAS As ICHCSII_DAS
Private mrsReport As ADODB.Recordset

Private Function CreateRSFromFrameworkFlexGrid(FlexGrid As VSFlexGrid) As Recordset

'Creates a recordset from a flexgrid
'The first row contains the column names

Dim oRs As Recordset
Dim iRowCounter As Long
Dim iColCounter As Long

    Set oRs = New Recordset
    oRs.CursorLocation = adUseClient

    With FlexGrid
        
        'Add the fields (cols)
        oRs.Fields.Append "LINE_NUM", adBSTR
        oRs.Fields.Append "TOOTH_NUM", adBSTR
        oRs.Fields.Append "REST_LOCATION", adBSTR
        oRs.Fields.Append "GUIDE_PLANE_LOCATION", adBSTR
        oRs.Fields.Append "RECIPROCATION_LOCATION", adBSTR
        oRs.Fields.Append "RETENTIVE_ARM_TYPE", adBSTR
        oRs.Fields.Append "GAUGE", adBSTR
        oRs.Fields.Append "UNDERCUT_LOCATION", adBSTR
        oRs.Fields.Append "UNDERCUT_AMOUNT", adBSTR
        
        oRs.Open
        
        'Add the data
        For iRowCounter = 1 To .Rows - 1
            If .Cell(flexcpText, iRowCounter, 1) <> "" Then
                oRs.AddNew
                For iColCounter = 0 To .Cols - 2
                    oRs.Fields(iColCounter) = .Cell(flexcpTextDisplay, iRowCounter, iColCounter)
                Next
                oRs.Update
            End If
        Next
        
    End With
    
    Set CreateRSFromFrameworkFlexGrid = oRs
End Function

Private Function CreateRSFromPonticsFlexGrid(FlexGrid As VSFlexGrid) As Recordset

'Creates a recordset from a flexgrid
'The first row contains the column names

Dim oRs As Recordset
Dim iRowCounter As Long
Dim iColCounter As Long

    Set oRs = New Recordset
    oRs.CursorLocation = adUseClient

    With FlexGrid
        
        'Add the fields (cols)
        oRs.Fields.Append "LINE_NUM", adBSTR
        oRs.Fields.Append "TOOTH_NUM", adBSTR
        oRs.Fields.Append "MATERIAL", adBSTR
        oRs.Fields.Append "GUIDE", adBSTR
        oRs.Fields.Append "SHADE", adBSTR
        oRs.Fields.Append "MOLD", adBSTR
        oRs.Fields.Append "RETENTION", adBSTR
        
        oRs.Open
        
        'Add the data
        For iRowCounter = 1 To .Rows - 1
            If .Cell(flexcpText, iRowCounter, 1) <> "" Then
                oRs.AddNew
                For iColCounter = 0 To 6
                    oRs.Fields(iColCounter) = .Cell(flexcpTextDisplay, iRowCounter, iColCounter)
                Next
                oRs.Update
            End If
        Next
        
    End With
    
    Set CreateRSFromPonticsFlexGrid = oRs
End Function
Private Sub Class_Initialize()
'   'Set moDAS = New CHCSII_DAS
'   If moDAS Is Nothing Then
'      If gobjShared.UseLocalDB Then
'         Set moDAS = New GEMS_DAS
'      Else
'         Set moDAS = New CHCSII_DAS
'      End If
'   End If
End Sub

Private Sub Class_Terminate()
'   Set moDAS = Nothing
   Set mrsReport = Nothing
End Sub

Public Function PrintReport(ByVal vSpecCategory As String, _
                              Optional ByVal vbPreview As Boolean = True) As Boolean
                              
   Dim oForm As frmLabReport
   Dim oCrownReport As CrownReport
   Dim oFPDReport As FPDReport
   Dim oRPDReport As RPDReport
   Dim sSql As String
   Dim oRPDFrs As Recordset
   Dim oRPDPrs As Recordset
   Dim oSubReport As Report
   
   'Variables for Report Formula
   
   '----- LabCase Information ---
   Dim sLabCaseNo As String
   Dim sToothLocation As String
   Dim sADACode As String
   
   '----- Crown & FPD ------
   Dim sTitle As String
   Dim sFrameworkMaterial As String
   Dim sVeneerMaterial As String
   Dim sGuide As String
   Dim sShade As String
   Dim sOcclusalSurfaceMaterial As String
   Dim sOcclusalScheme As String
   Dim sMargin As String
   Dim sMProximalContact As String
   Dim sDProximalContact As String
   Dim sGuidePlane As String
   Dim sRest1 As String
   Dim sRest2 As String
   Dim sUndercutLocation As String
   Dim sUndercutAmount As String
   Dim sAttachment As String
   Dim sAttachmentLocation As String
   Dim sMetalFinish As String
   Dim sVeneerFinish As String
   Dim sCheckAt As String
   Dim sReturnAt As String
   Dim sViewNotes As String
   
   '------- FPD -------
   Dim sMConnector As String
   Dim sDConnector As String
   Dim sPonticType As String
   Dim sDistanceFromGingiva As String
   Dim sTissueContactArea As String
   Dim sTissueContactMaterial As String
      
   'error handling
   On Error GoTo ErrHandler
   
   PrintReport = False
   'Load oForm
   Screen.MousePointer = vbHourglass
   Set oForm = New frmLabReport
   
   Select Case vSpecCategory
      Case gcCrown
         sTitle = "Crown Specifications"
         Set oCrownReport = New CrownReport
         
         'Assign object variables to local variables
         With frmLabs
            sLabCaseNo = .txtCrownLabCaseNo
            sToothLocation = .txtCrownToothLocation
            sADACode = .txtCrownADACode
            'KDunne 9/11/2002
            'SCR 8897
            sViewNotes = .txtRemarks(1).Text
            'Formula's field limitation
                     
         End With
         
         With frmLabs.cboCrownMaterial
            sFrameworkMaterial = .Item(0).Text
            sVeneerMaterial = .Item(1).Text
            sGuide = .Item(2).Text
            sShade = .Item(3).Text
            sOcclusalSurfaceMaterial = .Item(4).Text
            sOcclusalScheme = .Item(5).Text
            sMargin = .Item(6).Text
            sMProximalContact = .Item(7).Text
            sDProximalContact = .Item(8).Text
            sGuidePlane = .Item(9).Text
            sRest1 = .Item(10).Text
            sRest2 = .Item(11).Text
            sUndercutLocation = .Item(12).Text
            sUndercutAmount = .Item(13).Text
            sAttachment = .Item(14).Text
            sAttachmentLocation = .Item(15).Text
            sMetalFinish = .Item(16).Text
            sVeneerFinish = .Item(17).Text
            sCheckAt = .Item(18).Text
            sReturnAt = .Item(19).Text
         End With
            
         'Set HeaderInfo
         SetSpecReportHeader oCrownReport, sLabCaseNo, sToothLocation, sADACode
         'Set Formulas
         With oCrownReport
            'Report title
            .fldTitle.Field.Text = "'" & sTitle & "'"
                        
            '.fldFrameworkMaterial.Field.Text = "'" & sFrameworkMaterial & "'"
            .fldFrameworkMaterial.SetText sFrameworkMaterial
            
            '.fldVeneerMaterial.Field.Text = "'" & sVeneerMaterial & "'"
            .fldVeneerMaterial.SetText sVeneerMaterial
            
            '.fldGuide.Field.Text = "'" & sGuide & "'"
            .fldGuide.SetText sGuide
            
            '.fldShade.Field.Text = "'" & sShade & "'"
            .fldShade.SetText sShade
            
            '.fldOcclusalSurfaceMaterial.Field.Text = "'" & sOcclusalSurfaceMaterial & "'"
            .fldOcclusalSurfaceMaterial.SetText sOcclusalSurfaceMaterial
            
            '.fldOcclusalScheme.Field.Text = "'" & sOcclusalScheme & "'"
            .fldOcclusalScheme.SetText sOcclusalScheme
            
            '.fldMargin.Field.Text = "'" & sMargin & "'"
            .fldMargin.SetText sMargin
            
            '.fldMProximalContact.Field.Text = "'" & sMProximalContact & "'"
            .fldMProximalContact.SetText sMProximalContact
            
            '.fldDProximalContact.Field.Text = "'" & sDProximalContact & "'"
            .fldDProximalContact.SetText sDProximalContact
            
            '.fldGuidePlane.Field.Text = "'" & sGuidePlane & "'"
            .fldGuidePlane.SetText sGuidePlane
            
            '.fldRest1.Field.Text = "'" & sRest1 & "'"
            .fldRest1.SetText sRest1
            
            '.fldRest2.Field.Text = "'" & sRest2 & "'"
            .fldRest2.SetText sRest2
            
            '.fldUndercutLocation.Field.Text = "'" & sUndercutLocation & "'"
            .fldUndercutLocation.SetText sUndercutLocation
            
            '.fldUndercutAmount.Field.Text = "'" & sUndercutAmount & "'"
            .fldUndercutAmount.SetText sUndercutAmount
            
            '.fldAttachment.Field.Text = "'" & sAttachment & "'"
            .fldAttachment.SetText sAttachment
            
            '.fldAttachmentLocation.Field.Text = "'" & sAttachmentLocation & "'"
            .fldAttachmentLocation.SetText sAttachmentLocation
            
            '.fldMetalFinish.Field.Text = "'" & sMetalFinish & "'"
            .fldMetalFinish.SetText sMetalFinish
            
            '.fldVeneerFinish.Field.Text = "'" & sVeneerFinish & "'"
            .fldVeneerFinish.SetText sVeneerFinish
            
            '.fldCheckAt.Field.Text = "'" & sCheckAt & "'"
            .fldCheckAt.SetText sCheckAt
            
            '.fldReturnAt.Field.Text = "'" & sReturnAt & "'"
            .fldReturnAt.SetText sReturnAt
            
            'KDunne 9/11/2002
            'SCR 8897
            .fldViewNotes.SetText sViewNotes
            
         End With
   
         If vbPreview Then
            'Display report
            With oForm.CRViewer1
               .ReportSource = oCrownReport
               .DisplayGroupTree = False
               .EnableGroupTree = False
               .DisplayTabs = False
               .EnableDrillDown = False
               .EnableRefreshButton = False
               DoEvents
               .ViewReport
               DoEvents
               .Zoom 2
            End With
            
            'SCR 33933 - Don't display crystal reports as a modal window.
            'cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
            oForm.Show 'vbModal
            'cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
            
         Else
            'Print report
            oCrownReport.PrintOut False
         End If
         
      Case gcFPD
         sTitle = "FPD Specifications"
         Set oFPDReport = New FPDReport
         
         'Assign object variables to local variables
         With frmLabs
            sLabCaseNo = .txtFPDCaseNo
            sToothLocation = .txtFPDToothLocation
            sADACode = .txtFPDADACode
            'KDunne 9/11/2002
            'SCR 8897
            sViewNotes = .txtRemarks(2).Text
            'Formula's field limitation
            
         End With
         
         With frmLabs.cboFPDMaterial
            sFrameworkMaterial = .Item(0).Text
            sVeneerMaterial = .Item(1).Text
            sGuide = .Item(2).Text
            sShade = .Item(3).Text
            sOcclusalSurfaceMaterial = .Item(4).Text
            sOcclusalScheme = .Item(5).Text
            sMargin = .Item(6).Text
            sMProximalContact = .Item(7).Text
            sDProximalContact = .Item(8).Text
            sGuidePlane = .Item(9).Text
            sRest1 = .Item(10).Text
            sRest2 = .Item(11).Text
            sUndercutLocation = .Item(12).Text
            sUndercutAmount = .Item(13).Text
            sAttachment = .Item(14).Text
            sAttachmentLocation = .Item(15).Text
            sMetalFinish = .Item(16).Text
            sVeneerFinish = .Item(17).Text
            sMConnector = .Item(18).Text
            sDConnector = .Item(19).Text
            sPonticType = .Item(20).Text
            sDistanceFromGingiva = .Item(21).Text
            sTissueContactArea = .Item(22).Text
            sTissueContactMaterial = .Item(23).Text
            sCheckAt = .Item(24).Text
            sReturnAt = .Item(25).Text
         End With
            
         'Set Report Header
         SetSpecReportHeader oFPDReport, sLabCaseNo, sToothLocation, sADACode
         'Set Formulas
         With oFPDReport
            'Report title
            .fldTitle.Field.Text = "'" & sTitle & "'"
            
            '.fldFrameworkMaterial.Field.Text = "'" & sFrameworkMaterial & "'"
            .fldFrameworkMaterial.SetText sFrameworkMaterial
            
            '.fldVeneerMaterial.Field.Text = "'" & sVeneerMaterial & "'"
            .fldVeneerMaterial.SetText sVeneerMaterial
            
            '.fldGuide.Field.Text = "'" & sGuide & "'"
            .fldGuide.SetText sGuide
            
            '.fldShade.Field.Text = "'" & sShade & "'"
            .fldShade.SetText sShade
            
            '.fldOcclusalSurfaceMaterial.Field.Text = "'" & sOcclusalSurfaceMaterial & "'"
            .fldOcclusalSurfaceMaterial.SetText sOcclusalSurfaceMaterial
            
            '.fldOcclusalSurfaceScheme.Field.Text = "'" & sOcclusalScheme & "'"
            .fldOcclusalSurfaceScheme.SetText sOcclusalScheme
            
            '.fldMargin.Field.Text = "'" & sMargin & "'"
            .fldMargin.SetText sMargin
            
            '.fldMProximalContact.Field.Text = "'" & sMProximalContact & "'"
            .fldMProximalContact.SetText sMProximalContact
            
            '.fldDProximalContact.Field.Text = "'" & sDProximalContact & "'"
            .fldDProximalContact.SetText sDProximalContact
            
            '.fldGuidePlane.Field.Text = "'" & sGuidePlane & "'"
            .fldGuidePlane.SetText sGuidePlane
            
            '.fldRest1.Field.Text = "'" & sRest1 & "'"
            .fldRest1.SetText sRest1
            
            '.fldRest2.Field.Text = "'" & sRest2 & "'"
            .fldRest2.SetText sRest2
            
            '.fldUndercutLocation.Field.Text = "'" & sUndercutLocation & "'"
            .fldUndercutLocation.SetText sUndercutLocation
            
            '.fldUndercutAmount.Field.Text = "'" & sUndercutAmount & "'"
            .fldUndercutAmount.SetText sUndercutAmount
            
            '.fldAttachment.Field.Text = "'" & sAttachment & "'"
            .fldAttachment.SetText sAttachment
            
            '.fldAttachmentLocation.Field.Text = "'" & sAttachmentLocation & "'"
            .fldAttachmentLocation.SetText sAttachmentLocation
            
            '.fldMetalFinish.Field.Text = "'" & sMetalFinish & "'"
            .fldMetalFinish.SetText sMetalFinish
            
            '.fldVeneerFinish.Field.Text = "'" & sVeneerFinish & "'"
            .fldVeneerFinish.SetText sVeneerFinish
            
            '.fldMConnector.Field.Text = "'" & sMConnector & "'"
            .fldMConnector.SetText sMConnector
            
            '.fldDConnector.Field.Text = "'" & sDConnector & "'"
            .fldDConnector.SetText sDConnector
            
            '.fldPonticType.Field.Text = "'" & sPonticType & "'"
            .fldPonticType.SetText sPonticType
            
            '.fldDistanceFromGingiva.Field.Text = "'" & sDistanceFromGingiva & "'"
            .fldDistanceFromGingiva.SetText sDistanceFromGingiva
            
            '.fldTissueContactArea.Field.Text = "'" & sTissueContactArea & "'"
            .fldTissueContactArea.SetText sTissueContactArea
            
            '.fldTissueContactMaterial.Field.Text = "'" & sTissueContactMaterial & "'"
            .fldTissueContactMaterial.SetText sTissueContactMaterial
            
            '.fldCheckAt.Field.Text = "'" & sCheckAt & "'"
            .fldCheckAt.SetText sCheckAt
            
            '.fldReturnAt.Field.Text = "'" & sReturnAt & "'"
            .fldReturnAt.SetText sReturnAt
            
            'KDunne 9/11/2002
            'SCR 8897
            .fldViewNotes.SetText sViewNotes
            
         End With
   
         If vbPreview Then
            'Display report
            With oForm.CRViewer1
               .ReportSource = oFPDReport
               .DisplayGroupTree = False
               .EnableGroupTree = False
               .DisplayTabs = False
               .EnableDrillDown = False
               .EnableRefreshButton = False
               DoEvents
               .ViewReport
               DoEvents
               .Zoom 2
            End With


            'SCR 33933 - Don't display crystal reports as a modal window.
            'cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
            oForm.Show 'vbModal
            'cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
            
            cwobjCWComm.Message cwiREFRESH_DATA, "", NCID_DDSA_APPLICATION, NCID_DENTAL_LABS_APPLICATION, False

            
            'frmLabs.txtFPDNewNotes.Refresh
            
         Else
            'Print report
            oFPDReport.PrintOut False
         End If
      'DoEvents
      
      Case gcRPD
         sTitle = "RPD Specifications"
         Set oRPDReport = New RPDReport
         
         'Assign object variables to local variables
         With frmLabs
            sLabCaseNo = .txtRPDCaseNo
            'sToothLocation = .txtFPDToothLocation
            sADACode = .txtRPDADACode
            'KDunne 9/11/2002
            'SCR 8897
            sViewNotes = .txtRemarks(3).Text
            'Formula's field limitation
            
         End With
         
         With frmLabs.cboRPDMaterial
            sFrameworkMaterial = .Item(0).Text
            sMConnector = .Item(1).Text
            sOcclusalSurfaceMaterial = .Item(2).Text
            sOcclusalScheme = .Item(3).Text
            sShade = .Item(3).Text
            
            sCheckAt = .Item(8).Text
            sReturnAt = .Item(9).Text
         End With
            
         'Set Report Header
         SetSpecReportHeader oRPDReport, sLabCaseNo, "", sADACode
         'Set Formulas
         With oRPDReport
            'Report title
            .fldTitle.Field.Text = "'" & sTitle & "'"
            
            'RPD Framework
            .fldFrameworkMaterial.SetText sFrameworkMaterial
            .fldMajorConnector.SetText sMConnector
            
            'Add the framework grid
            Set oRPDFrs = CreateRSFromFrameworkFlexGrid(frmLabs.VSFlexGridRPDFramework)
            'Open the subreport
            Set oSubReport = .OpenSubreport("SubRptRPD")
            oSubReport.Database.SetDataSource oRPDFrs
            
            'RPD Pontics
            .fldOcclusalSurfaceMaterial.SetText sOcclusalSurfaceMaterial
            .fldOcclusalScheme.SetText sOcclusalScheme
            .fldDentureBaseMaterial.SetText frmLabs.cboRPDMaterial.Item(4).Text
            .fldDentureBaseRetention.SetText frmLabs.cboRPDMaterial.Item(5).Text
            .fldResinType.SetText frmLabs.cboRPDMaterial.Item(6).Text
            .fldResinShade.SetText frmLabs.cboRPDMaterial.Item(7).Text
            
            If frmLabs.chkRPDMaterial.Item(0).Value = 1 Then
                .fldPreserveFacebow.SetText "Yes"
            End If
            
            If frmLabs.chkRPDMaterial.Item(1).Value = 1 Then
                .fldRemountCastsAndMountMaxillary.SetText "Yes"
            End If
            
            If frmLabs.chkRPDMaterial.Item(2).Value = 1 Then
                .fldLabRemount.SetText "Yes"
            End If
            
            If frmLabs.chkRPDMaterial.Item(3).Value = 1 Then
                .fldLabOcclusalAdjustment.SetText "Yes"
            End If
            
            If frmLabs.chkRPDMaterial.Item(4).Value = 1 Then
                .fldCustomCharacterization.SetText "Yes"
            End If
            
            If frmLabs.chkRPDMaterial.Item(5).Value = 1 Then
                .fldStippling.SetText "Yes"
            End If
            
            'Add the pontics grid
            Set oRPDPrs = CreateRSFromPonticsFlexGrid(frmLabs.VSFlexGridRPDPontics)
            'Open the subreport
            Set oSubReport = .OpenSubreport("SubRptPontics")
            oSubReport.Database.SetDataSource oRPDPrs
            
            .fldCheckAt.SetText sCheckAt
            .fldReturnAt.SetText sReturnAt
            'KDunne 9/11/2002
            'SCR 8897
            .fldViewNotes.SetText sViewNotes
            
         End With
   
         If vbPreview Then
            'Display report
            With oForm.CRViewer1
               .ReportSource = oRPDReport
               .DisplayGroupTree = False
               .EnableGroupTree = False
               .DisplayTabs = False
               .EnableDrillDown = False
               .EnableRefreshButton = False
               DoEvents
               .ViewReport
               DoEvents
               .Zoom 2
            End With

            'SCR 33933 - Don't display crystal reports as a modal window.
            'cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
            oForm.Show 'vbModal
            'cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
            
            cwobjCWComm.Message cwiREFRESH_DATA, "", NCID_DDSA_APPLICATION, NCID_DENTAL_LABS_APPLICATION, False

            
            'frmLabs.txtFPDNewNotes.Refresh
            
         Else
            'Print report
            oRPDReport.PrintOut False
         End If
         
   End Select
      
   PrintReport = True
   
   GoTo CleanUp
   
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabReport.PrintReport", "Labs", vbCritical)
CleanUp:
   Set oCrownReport = Nothing
   Set oFPDReport = Nothing
   Set oRPDReport = Nothing
   Set oForm = Nothing
   Screen.MousePointer = vbDefault
   Exit Function
   Resume
End Function
'*************************************************************************
'Function Name: FirstLastMiName
'Description: Convert full name as Last, First, Middle initial format
'*************************************************************************
Private Function FirstLastMiName(ByVal pFullName As String) As String
    Dim vName As Variant
    
    vName = Split(pFullName, " ", , vbTextCompare)
    If UBound(vName, 1) < 2 Then
        'Return original full name
        FirstLastMiName = pFullName
    Else
        'Return formatted name
        FirstLastMiName = vName(0) & " " & vName(1) & " " & Left(vName(2), 1)
    End If

End Function

'***************************************************************************
'Function:FormatSSN
'Author:KDunne
'Date:8/14/2002
'Purpose: Formats Social Security Number
'****************************************************************************
Private Function FormatSSN(ByVal strSSN As String) As String
    Dim strReturn As String
    
    If IsNumeric(strSSN) Then
        strReturn = Format(strSSN, "00-00-0000")
    Else
        strReturn = strSSN
    End If
    
    FormatSSN = strReturn
End Function

Private Sub SetSpecReportHeader(ByRef objReport As Object, ByVal sLabCaseNo As String, _
        ByVal sToothLocation As String, ByVal sADACode As String)

Dim patientAge As Integer
Dim sCalcAge As String
Dim Age As Integer
    
    'SCR Description: Calculate Patient Age from cwshared method which is what the patient ID bar is using so Dental will match Patient ID bar all the time.
     If (cwobjPatient.DateOfBirth = DateValue("31-Dec-9999")) Or cwobjPatient.DateOfBirth = 0 Then
        patientAge = 20
     Else
        ' Get the patient demographic data available
        ' Age in days = difference between birthdate and encounter time
        ' Need age in minutes
        sCalcAge = gobjShared.CalcAge(cwobjPatient.DateOfBirth)
        
        If InStr(sCalcAge, "mo") > 0 Then
            If CInt(Left(sCalcAge, Len(sCalcAge) - 2)) < 12 Then
                patientAge = 0
            Else
                patientAge = CInt(Left(sCalcAge, Len(sCalcAge) - 2) / 12)
                
            End If
        Else
        
            patientAge = CInt(Left(sCalcAge, Len(sCalcAge) - 2))
        End If
    End If
    
    With objReport
        'Report fields
        .fldLabCaseNo.SetText sLabCaseNo
        '.fldAge.SetText CStr(Fix((Date - cwobjPatient.DateOfBirth) / 365))
        .fldAge.SetText patientAge
        .fldGrade.SetText cwobjPatient.Sponsor.Rank
        .fldPatientName.SetText FirstLastMiName(cwobjPatient.Name)
        .fldSSN.SetText FormatSSN(cwobjPatient.SSN)
        .fldTreatmentFacility.SetText guCurEncInfo.ClinicName
        If sToothLocation <> "" Then
            .fldToothLocation.SetText sToothLocation
        End If
        .fldADACode.SetText sADACode
    End With
End Sub
