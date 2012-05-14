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
Attribute VB_Name = "FormQuestionnaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormQuestionnaire.cls"
Private moPatientDemographics As PatientDemographics
Private moSponsorRec As ADODB.Recordset
Private moPatientRec As ADODB.Recordset

Public Function GetHeader() As String

    Dim sHeader As String
    
    sHeader = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f16\froman\fcharset238\fprq2 Times New Roman CE;}{\f17\froman\fcharset204\fprq2 Times New Roman Cyr;}" & _
    "{\f19\froman\fcharset161\fprq2 Times New Roman Greek;}{\f20\froman\fcharset162\fprq2 Times New Roman Tur;}{\f21\froman\fcharset186\fprq2 Times New Roman Baltic;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;" & _
    "\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;" & _
    "\red128\green128\blue128;\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\cgrid \snext0 Normal;}{\s1\keepn\widctlpar\adjustright \b\fs20\cgrid \sbasedon0 \snext0 heading 1;}{\s2\keepn\widctlpar\adjustright \b\fs18\cgrid " & _
    "\sbasedon0 \snext0 heading 2;}{\*\cs10 \additive Default Paragraph Font;}}{\info{\title Patient:}{\author webb}{\operator webb}{\creatim\yr2000\mo6\dy12\hr17\min50}{\revtim\yr2000\mo6\dy13\hr14\min25}{\version3}{\edmins1}{\nofpages1}{\nofwords25}" & _
    "{\nofchars145}{\*\company usgov}{\nofcharsws0}{\vern89}}\margl576\margr576\margt576\margb576 \widowctrl\ftnbj\aenddoc\hyphcaps0\formshade\viewkind1\viewscale100\pgbrdrhead\pgbrdrfoot \fet0\sectd \linex0\endnhere\sectdefaultcl {\*\pnseclvl1" & _
    "\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5" & _
    "\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang" & _
    "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\trowd \trgaph108\trleft-108\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \trbrdrh\brdrs\brdrw10 \trbrdrv" & _
    "\brdrs\brdrw10 \clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx160\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx10710\pard\plain " & _
    "\s1\qc\keepn\widctlpar\intbl\outlinelevel0\adjustright \b\fs20\cgrid {\cell QUESTIONNAIRE/TEST RESPONSE\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\row }\pard \widctlpar\adjustright {" & _
    "\par }}"
        
    GetHeader = sHeader

End Function
Public Function Build(sRTF As String, sTitle As String, bPreview As Boolean, _
                           Optional iNumberOfCopies As Integer = 1) As Boolean

    On Error GoTo ErrHandler
    Build = False
    
    Dim i As Integer
    Dim iPrintCount As Integer
    Dim lPageFrom As Long
    Dim lPageTo As Long
    
    Dim FormatDTS As String
    Dim iPages As Integer
    Dim CurPage As Integer
    Dim sHeader1 As String
    Dim sFooter As String
    Dim sBody As String
    Dim sUpdatedBody As String
    
    Dim xxx As String
    Dim sTemp As String
    Dim sStatus As String
    Dim sRelToPat As String
    Dim nCurrentEnctrStatus As Integer
    Dim sMedicalRecordNumber As String
    Dim sHeader As String
    Dim sContextNCID As String
    Dim sSponsorUnitNumber As String
    
    #If coiDebug Then
         Dim DR As DLog.DebugRec
         gCOI_DL.DebugLog "COI-MOP#2-1-1", "", "", "Begin", DR
    #End If

    '-- SF600 HEADER    '------------------------------------------------------------------------------
    sHeader = GetHeader()
    frmSF600.txHeader.PageWidth = 0 '- reset fir autosize
    frmSF600.txHeader.PageHeight = 0 '- reset fir autosize
    frmSF600.txHeader.Text = ""
    frmSF600.txHeader.Height = 1400
    frmSF600.txHeader.PageMarginT = 450
    frmSF600.txHeader.PageMarginL = 600

    '--- SCR 17046
    sBody = Replace(sRTF, "\b\qc", "\b\ql")

    '-- SF600 BODY    '------------------------------------------------------------------------------
    frmSF600.txBody.PageWidth = 0 '- reset fir autosize
    frmSF600.txBody.PageHeight = 0 '- reset fir autosize
    frmSF600.txBody.AutoExpand = True
    frmSF600.txBody.Text = ""
    frmSF600.txBody.Height = 500
    
    frmSF600.txBody.PageMarginT = 950
    frmSF600.txBody.PageMarginL = 1440
    frmSF600.txBody.PageMarginB = 3220
    frmSF600.txBody.PageHeight = 15748
    frmSF600.txBody.PageWidth = 12240
    frmSF600.txBody.RTFSelText = sRTF
    
    '-- SF600 FOOTER    '------------------------------------------------------------------------------
    frmSF600.txFooter.PageWidth = 0 '- reset fir autosize
    frmSF600.txFooter.PageHeight = 0 '- reset fir autosize
    frmSF600.txFooter.Text = ""
    frmSF600.txFooter.AutoExpand = True
    frmSF600.txFooter.RTFSelText = sFooter
    frmSF600.txFooter.PageMarginT = 15000
    frmSF600.txFooter.PageMarginL = 600

    iPages = frmSF600.txBody.CurrentPages


    '-- PRINT PREVIEW    '-------------------------------------------------------------------------------
    If gobjshared.UseSQLServer And Not gobjshared.IsAppMode(modeCHCSI_GUI) Then
        frmSF600.txWaterMark.ObjectInsertAsChar 0, gobjshared.ImagesPath & "\TRAINING.TIF", -1, 250, 250, &H2&, 1
        DoEvents
    End If
    
    If bPreview Then

        '-- Give Form information to print the SF600 by calling this procedure
        Set frmPreviewSF600.oCalledBy2 = Me
        frmPreviewSF600.Caption = "Questionnaire Response Print Preview"
        'frmPreviewSF600.sFacilityNCID = sFacilityNCID
        'frmPreviewSF600.sUnitNumber = sPatientUnitNumber
        frmPreviewSF600.sRTF = sRTF

        frmPreviewSF600.PicPreview(0).Height = 15840
        frmPreviewSF600.VSViewPort1.VirtualHeight = 15840
        frmPreviewSF600.PicPreview(0).Width = 13000
        frmPreviewSF600.VSViewPort1.VirtualWidth = 13000
        frmPreviewSF600.VSViewPort1.VirtualTop = 0
        frmPreviewSF600.VSViewPort1.SetVirtualExtent
        frmPreviewSF600.VSViewPort1.Track = True
        frmPreviewSF600.VSViewPort1.MouseScroll = True

        For i = 0 To iPages - 1
            If i > 0 Then
                Load frmPreviewSF600.PicPreview(i)
                SetParent frmPreviewSF600.PicPreview(i), frmPreviewSF600.VSViewPort1.hwnd
                frmPreviewSF600.PicPreview(i).Move frmPreviewSF600.PicPreview(0).left, frmPreviewSF600.PicPreview(0).top, frmPreviewSF600.PicPreview(0).Width, frmPreviewSF600.PicPreview(0).Height
            End If

            frmSF600.txHeader.Text = ""
            frmSF600.txHeader.RTFSelText = sHeader
            frmSF600.txHeader.ZOrder 1
            frmSF600.txHeader.PrintDevice = frmPreviewSF600.PicPreview(i).hDC
            frmSF600.txHeader.PrintPage 1

            frmSF600.txBody.PrintDevice = frmPreviewSF600.PicPreview(i).hDC
            frmSF600.txBody.PrintPage i + 1

            frmSF600.txFooter.ZOrder 1
            frmSF600.txFooter.Text = ""
            frmSF600.txFooter.RTFSelText = GetFooter
            frmSF600.txFooter.PrintDevice = frmPreviewSF600.PicPreview(i).hDC
            frmSF600.txFooter.PrintPage 1

        Next i

        #If coiDebug Then
             gCOI_DL.DebugLog "COI-MOP#2-1-1", "", "", "End", DR
        #End If

        frmPreviewSF600.Show vbModal

'-- PRINT
'-------------------------------------------------------------------------------
    Else
        If Printers.Count = 0 Then
           MsgBxARMd "No printers are installed. To install a printer, point to Settings on the Windows Start menu, " & _
                "click Printers, and then double-click Add Printer. Follow the instructions in the wizard.", vbOKOnly + vbExclamation
           Exit Function
        End If
        frmPrint.MaxPageTo = iPages
        frmPrint.Show vbModal
        If frmPrint.PrintOperationCancelled Then
           Unload frmPrint
           GoTo CleanUp
        End If
        iNumberOfCopies = frmPrint.NumberOfCopies
        If frmPrint.PrintAll Then
           lPageFrom = 0
           lPageTo = iPages - 1
        Else
           lPageFrom = frmPrint.PageFrom - 1
           lPageTo = frmPrint.PageTo - 1
        End If
        Unload frmPrint
        
        For iPrintCount = 1 To iNumberOfCopies
            For i = lPageFrom To lPageTo
                Printer.Print
                
                If gobjshared.UseSQLServer Then
                    frmSF600.txWaterMark.PrintDevice = Printer.hDC
                    frmSF600.txWaterMark.PrintPage 1
                End If
    
                '--
                frmSF600.txHeader.Text = ""
                frmSF600.txHeader.RTFSelText = sHeader
                frmSF600.txHeader.ZOrder 1
                frmSF600.txHeader.PrintDevice = Printer.hDC
                frmSF600.txHeader.PrintPage 1
    
                '-- Print body
                frmSF600.txBody.PrintDevice = Printer.hDC
                frmSF600.txBody.PrintPage i + 1
    
                '-- print footer
                frmSF600.txFooter.Text = ""
                frmSF600.txFooter.RTFSelText = GetFooter
                frmSF600.txFooter.ZOrder 1
                frmSF600.txFooter.PrintDevice = Printer.hDC
                frmSF600.txFooter.PrintPage 1
                            
                Printer.NewPage
            Next i
        Next iPrintCount


        Printer.EndDoc
        Build = True

        #If coiDebug Then
             gCOI_DL.DebugLog "COI-MOP#2-1-1", "", "", "End", DR
        #End If

        GoTo CleanUp
    End If     '--  If bPreview Then

    If PrintPreviewCancel = "Cancel" Then Build = True
    GoTo CleanUp
    
ErrHandler:
  If Err = 480 Then
     Resume Next
  ElseIf Err = 482 Then
    If Printers.Count = 0 Then
       MsgBxARMd "No printers are installed. To install a printer, point to Settings on the Windows Start menu, " & _
                    "click Printers, and then double-click Add Printer. Follow the instructions in the wizard.", vbOKOnly + vbExclamation
    End If
  Else
     gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".Build", App.EXEName, vbExclamation
  End If

CleanUp:
    Unload frmPreviewSF600
    Set frmPreviewSF600 = Nothing
    Unload frmSF600
    Set frmSF600 = Nothing
End Function

Public Function GetFooter(Optional sCurrentPage As String, _
                               Optional sTotalPages As String) As String

    Dim sFooter As String
    
    GetFooter = ""
    
    sFooter = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033 {\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
            & "{\colortbl;\red0\green0\blue0;} \pard\plain \widctlpar\adjustright \cgrid"
    
    sFooter = sFooter & " \fs16 THIS INFORMATION IS PROTECTED BY THE PRIVACY ACT OF 1974 (PL-93-579). UNAUTHORIZED ACCESS " & _
    " TO THIS INFORMATION IS A VIOLATION OF FEDERAL LAW. VIOLATORS WILL BE PROSECUTED.} \par }"
    
    GetFooter = sFooter

End Function

Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub

Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
    Set moPatientRec = Nothing
End Sub


