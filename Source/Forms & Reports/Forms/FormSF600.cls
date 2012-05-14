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
Attribute VB_Name = "FormSF600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormSF600.cls"
Private moPatientDemographics As PatientDemographics
Private moSponsorRec As ADODB.Recordset
Private moPatientRec As ADODB.Recordset

Public mbGreenSwitch As Boolean
Private Function FixImageScale(sInputRTF As String, lPageHeight As Long) As String

    Dim sTempRTF As String
    Dim sTempRTF2 As String
    Dim sEditedRTF As String
    Dim sPictSeg As String
    
    Dim lPosition As Long
    Dim lTextPosition As Long
    Dim lTextPositionEnd As Long
    Dim lScaleX As Long
    Dim lScaleY As Long
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lFactor As Long
    
    On Error GoTo ErrHandler
    FixImageScale = sInputRTF
    
    lPosition = InStr(1, sInputRTF, "\pict", vbTextCompare)
    If lPosition = 0 Then Exit Function
    sTempRTF = sInputRTF
    
    Do:
        lPosition = InStr(1, sTempRTF, "\pict", vbTextCompare)
        If lPosition = 0 Then
           sEditedRTF = sEditedRTF & sTempRTF
           Exit Do
        End If
        
        sEditedRTF = sEditedRTF & left(sTempRTF, lPosition - 1)
        sTempRTF = right(sTempRTF, Len(sTempRTF) - lPosition + 1)
        sPictSeg = left(sTempRTF, InStr(1, sTempRTF, " ", vbTextCompare) - 1)
        lTextPosition = InStr(1, sPictSeg, "\pich", vbTextCompare) + 5
        lTextPositionEnd = InStr(lTextPosition, sPictSeg, "\", vbTextCompare)
        lHeight = CLng(Mid(sPictSeg, lTextPosition, lTextPositionEnd - lTextPosition))
        
        lTextPosition = InStr(1, sPictSeg, "\picw", vbTextCompare) + 5
        lTextPositionEnd = InStr(lTextPosition, sPictSeg, "\", vbTextCompare)
        lWidth = CLng(Mid(sPictSeg, lTextPosition, lTextPositionEnd - lTextPosition))
        
        lTextPosition = InStr(1, sPictSeg, "\picscalex", vbTextCompare) + 10
        lTextPositionEnd = InStr(lTextPosition, sPictSeg, "\", vbTextCompare)
        lScaleX = CLng(Mid(sPictSeg, lTextPosition, lTextPositionEnd - lTextPosition))
        
        lTextPosition = InStr(1, sPictSeg, "\picscaley", vbTextCompare) + 10
        lScaleY = CLng(right(sPictSeg, Len(sPictSeg) - lTextPosition + 1))
        If lHeight > lPageHeight Then
           lFactor = (lPageHeight / lHeight) * 100
           sPictSeg = Replace(sPictSeg, "\picscalex" & CStr(lScaleX), "\picscalex" & CStr(lFactor))
           sPictSeg = Replace(sPictSeg, "\picscaley" & CStr(lScaleY), "\picscaley" & CStr(lFactor))
        End If
        
        sEditedRTF = sEditedRTF & sPictSeg
        sTempRTF = right(sTempRTF, Len(sTempRTF) - InStr(1, sTempRTF, " ", vbTextCompare) + 1)
        
    Loop
    
    FixImageScale = sEditedRTF
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".FixImageScale", App.EXEName, vbExclamation
    frmSF600.txBody.PageHeight = 0 'Prevent application from locking up if image rtf cannot be processed
    
End Function


Public Function GetSF600Footer(sUnitNumber As String, _
                               Optional sCurrentPage As String, _
                               Optional sTotalPages As String) As String

    Dim sFooter As String
    Dim sPatientName As String
    Dim b2ndNameLine As Boolean
    
    Dim sFMPSponsorSSN As String
    Dim sSponsorSSN As String
    Dim sDOB As String
    Dim sPCat As String
    Dim sMCStatus As String
    Dim sInsurance As String
    Dim sSponsorName As String
    Dim sRank As String
    Dim sUnit As String
    Dim sOutptRecRm As String
    Dim sPCM As String
    Dim sPCMGroup As String
    Dim sCI As String
    Dim sSex As String
    Dim sHomePhone As String
    Dim sWorkPhone As String
    Dim sCS As String
    Dim sWS As String
    Dim sTelPCM As String
    
    GetSF600Footer = ""
    
    sPatientName = moPatientRec.Fields.Item("full_Name").Value  ''GetPatientsName(sUnitNumber)
    If Len(sPatientName) > 14 Then
        b2ndNameLine = True
    Else
        b2ndNameLine = False
    End If
    
    sSponsorSSN = moPatientRec.Fields.Item("Sponsor_SSN").Value
    sFMPSponsorSSN = moPatientRec.Fields.Item("FMP").Value & "/" & moPatientRec.Fields.Item("Sponsor_SSN").Value 'GetFMPSSN(sUnitNumber)
    sDOB = Format(moPatientRec.Fields.Item("DOB").Value, gsDefaultFormat) 'GetDOB(sUnitNumber)
    sPCat = "" & moPatientRec.Fields.Item("Patient_Category").Value 'GetPatientCategory(sUnitNumber)
    sMCStatus = moPatientRec.Fields.Item("Alternative_Care").Value & "" 'GetMCStatus(sUnitNumber)
    
    sSponsorName = moSponsorRec.Fields.Item("full_Name").Value & "" 'GetSponsorName(sUnitNumber)
    sRank = moSponsorRec.Fields.Item("Rank").Value & "" 'GetSponsorRank(sUnitNumber)
    sUnit = moSponsorRec.Fields.Item("UIC").Value & "" 'GetSponsorUnit(sUnitNumber)
    sOutptRecRm = moPatientRec.Fields.Item("Records_Maintained_At").Value 'GetRecordLocation(sUnitNumber)
    
    If gobjshared.IsAppMode(modeCHCSI_GUI) Then
       sPCM = moPatientRec.Fields.Item("PCM_ID").Value
       sTelPCM = moPatientRec.Fields.Item("TELPCM").Value
    Else
        GetPCM sUnitNumber, sPCM, sTelPCM
    End If
    
    sPCMGroup = GetPCMGroup(sUnitNumber)
    sSex = moPatientRec.Fields.Item("Sex").Value 'GetSex(sUnitNumber)
    'sCI = GetCI(sUnitNumber)
    sHomePhone = moPatientRec.Fields.Item("Home_Phone").Value 'GetHomePhone(sUnitNumber)
    sWorkPhone = moPatientRec.Fields.Item("Work_Phone").Value 'GetWorkPhone(sUnitNumber)
    'moPatientRec.Fields.Item("Command_Security").Value
    If Not gobjshared.IsAppMode(modeCHCSI_GUI) Then sCS = GetCS(sUnitNumber, gsContextNCID)
    If Not gobjshared.IsAppMode(modeCHCSI_GUI) Then sWS = GetWS(sUnitNumber)
    
    'If cwobjCWComm.CmdLineSwitch("GREEN") <> "" Then MsgBox "Green"
    If mbGreenSwitch Then
        'New Footer as requested by 28668
        sFooter = ""
        sFooter = _
          "{\rtf1\ansi\ansicpg1252\uc0\deff1{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}" & _
          "{\f1\froman\fcharset0\fprq2 Times New Roman;}{\f2\froman\fcharset2\fprq2 Symbol;}}" & _
          "{\colortbl;\red0\green0\blue0;\red255\green255\blue255;}{\stylesheet{\s0\f1\fs20 Normal;}{\*\cs10\additive Default Paragraph Font;}{\s16\f0\fs24 [Normal];}{\s17\f1\fs20\b\sbasedon0\snext0 heading 1;}{\s18\f1\fs18\b\sbasedon0\snext0 heading 2;}}"

        If b2ndNameLine Then      'Push up the name field
            sFooter = sFooter & _
              "\deftab1134\paperw12240\paperh15840\margl432\margt432\margr432\margb432\trowd" & _
              "\trgaph72\trleft0\trpaddl72\trpaddt0\trpaddr72\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3" & _
              "\trpaddfb3\clvertalt\clbrdrt\brdrs\brdrw10\cellx9333\clvertalt\clbrdrt\brdrs\brdrw10" & _
              "\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320" & _
              "\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
              "\f1\fs20\b Name: $FOOTER_NAME \cell\cell\intbl\row\pard\trowd\trgaph0\trleft0\trpaddl0" & _
              "\trpaddt0\trpaddr0\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt\cellx810" & _
              "\clvertalt\cellx3420\clvertalt\cellx4050\clvertalt\cellx6210\clvertalt\cellx7380\clvertalt" & _
              "\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040" & _
              "\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16\cell\pard\s17\intbl\tx720" & _
              "\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080" & _
              "\plain\f1\fs20\b\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480" & _
              "\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16 Sex:\cell\pard\s18\intbl\tx720\tx1440\tx2160" & _
              "\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        Else
            sFooter = sFooter & _
              "\deftab1134\paperw12240\paperh15840\margl432\margt432\margr432\margb432\trowd\trgaph0" & _
              "\trleft0\trpaddl0\trpaddt0\trpaddr0\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt" & _
              "\clbrdrt\brdrs\brdrw10\cellx810\clvertalt\clbrdrt\brdrs\brdrw10\cellx3420\clvertalt\clbrdrt\brdrs" & _
              "\brdrw10\cellx4050\clvertalt\clbrdrt\brdrs\brdrw10\cellx6210\clvertalt\clbrdrt\brdrs\brdrw10\cellx7380" & _
              "\clvertalt\clbrdrt\brdrs\brdrw10\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880" & _
              "\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16  Name:\cell\pard\" & _
              "s17\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080" & _
              "\plain\f1\fs20  \plain" & _
              "\f1\fs20\b $FOOTER_NAME\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200" & _
              "\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16 Sex:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320" & _
              "\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        End If
        
        sFooter = Replace$(sFooter, "$FOOTER_NAME", sPatientName, 1)
        
        'Sex field
        sFooter = sFooter & _
            "\f1\fs16\b " & sSex & "\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920" & _
                "\tx8640\tx9360\tx10080\plain\f1\fs16 Sponsor/SSN:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        'Sponsor field
        sFooter = sFooter & _
            "\f1\fs16 " & sSponsorName & "/" & sSponsorSSN & "\cell\intbl\row\pard\trowd\trgaph36\trleft0\trpaddl36\trpaddt0\trpaddr36\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3" & _
                "\trpaddfb3\clvertalt\cellx810\clvertalt\cellx3420\clvertalt\cellx4050\clvertalt\cellx6210\clvertalt\cellx7380\clvertalt\cellx11070\pard" & _
                "\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1" & _
                "\fs16 FMP/SSN:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        'SMP/SSN field and Home No
        sFooter = sFooter & _
            "\f1\fs20\b " & sFMPSponsorSSN & "\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 Tel H:\cell " & sHomePhone & "\plain\f1\fs16\b\cell\plain\f1\fs16 Rank:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040" & _
                "\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        'Rank, DOB, Work No,
        sFooter = sFooter & _
            "\f1\fs16 " & sRank & "\cell\intbl\row\pard\trowd\trgaph36\trleft0\trpaddl36\trpaddt0\trpaddr36\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3" & _
                "\clvertalt\cellx810\clvertalt\cellx3420\clvertalt\cellx4050\clvertalt\cellx6210\clvertalt\cellx7380\clvertalt\cellx11070\pard\widctlpar\intbl" & _
                "\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16 DOB:\cell\pard\s18" & _
                "\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16\b " & sDOB & "\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 Tel W:\cell " & sWorkPhone & "\plain\f1\fs16\b\cell\plain\f1\fs16 Unit:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760" & _
                "\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        'Unit, PCAT, CS, REC_RM
        sFooter = sFooter & _
            "\f1\fs16 " & sUnit & "\cell\intbl\row\pard\trowd\trgaph36\trleft0\trpaddl36\trpaddt0\trpaddr36\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt" & _
                "\cellx810\clvertalt\cellx3420\clvertalt\cellx4050\clvertalt\cellx6210\clvertalt\cellx7380\clvertalt\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720" & _
                "\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16 PCat:\cell\pard\s18\intbl\tx720\tx1440\tx2160" & _
                "\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 " & sPCat & "\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 CS:\cell " & sCS & "\cell Outpt Rec. Rm:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 " & sOutptRecRm & "\cell\intbl\row\pard\trowd\trgaph36\trleft0\trpaddl36\trpaddt0\trpaddr36\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt" & _
                "\cellx810\clvertalt\cellx3420\clvertalt\cellx4050\clvertalt\cellx6210\clvertalt\cellx7380\clvertalt\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440" & _
                "\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16 MC Status:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600" & _
                "\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
          'MC_Status, WS, PCM, Insurance
        sFooter = sFooter & _
            "\f1\fs16 " & sMCStatus & "\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 SWS:\cell " & sWS & "\cell PCM:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 " & sPCM & "\cell\intbl\row\pard\trowd\trgaph36\trleft0\trpaddl36\trpaddt0\trpaddr36\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt\clbrdrb\brdrs" & _
                "\brdrw10\cellx810\clvertalt\clbrdrb\brdrs\brdrw10\cellx3420\clvertalt\clbrdrb\brdrs\brdrw10\cellx4050\clvertalt\clbrdrb\brdrs\brdrw10\cellx6210\clvertalt\clbrdrb\brdrs" & _
                "\brdrw10\cellx7380\clvertalt\clbrdrb\brdrs\brdrw10\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 Insurance:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
            "\f1\fs16 " & sInsurance & "\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16\cell\plain" & _
                "\f1\fs16\b\cell\plain\f1\fs16 Tel. PCM:\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain"
        'PCM No and the rest
        sFooter = sFooter & _
            "\f1\fs16 " & sTelPCM & "\cell\intbl\row\pard\trowd\trgaph144\trleft0\trpaddl144\trpaddt0\trpaddr144\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt\brdrs\brdrw10"
            '"\f1\fs16 " & sTelPCM & "\cell\intbl\row\pard\trowd\trgaph144\trleft0\trpaddl144\trpaddt0\trpaddr144\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt\clbrdrt\brdrs\brdrw10"
                '"\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs6\b\cell\intbl\row\pard"
        sFooter = sFooter & _
                "\cellx11070\pard\widctlpar\intbl\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\brdrnone\plain\f1\fs6\cell\intbl\row\pard" & _
                "\trowd\trgaph72\trleft0\trpaddl72\trpaddt0\trpaddr72\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt\cellx8586\clvertalt\cellx8820\clvertalt\cellx11070\pard\widctlpar" & _
                "\intbl\pard\intbl\qc\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs14\b CHRONOLOGICAL RECORD OF MEDICAL CARE\par\pard" & _
                "\intbl\qc\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs16 THIS INFORMATION IS PROTECTED BY THE PRIVACY ACT OF 1974 (PL-93-579). UNAUTHORIZED ACCESS \par TO THIS INFORMATION IS A VIOLATION OF FEDERAL LAW. VIOLATORS WILL BE PROSECUTED.\plain" & _
                "\f1\fs14\b\cell\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\cell\pard\s18\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480" & _
                "\tx7200\tx7920\tx8640\tx9360\tx10080\plain\f1\fs14\b STANDARD FORM 600 (REV. 5)\par\pard\intbl\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
                "\f1\fs14\b Prescribed by GSA and ICMR\par FIRMR (41 CFR) 201-45.505\cell\intbl\row\pard\trowd\trgaph144\trleft0\trpaddl144\trpaddt0\trpaddr144\trpaddb0\trpaddfl3\trpaddft3\trpaddfr3\trpaddfb3\clvertalt" & _
                "\cellx11070\pard\widctlpar\intbl\pard\intbl\qr\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\plain" & _
                "\f1\fs16\b Page " & sCurrentPage & " of " & sTotalPages & "\cell\intbl\row\pard\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\tx9360\tx10080\par }"
    Else
        sFooter = _
        "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f16\froman\fcharset238\fprq2 Times New Roman CE;}{\f17\froman\fcharset204\fprq2 Times New Roman Cyr;}" & _
        "{\f19\froman\fcharset161\fprq2 Times New Roman Greek;}{\f20\froman\fcharset162\fprq2 Times New Roman Tur;}{\f21\froman\fcharset186\fprq2 Times New Roman Baltic;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;" & _
        "\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;" & _
        "\red128\green128\blue128;\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\cgrid \snext0 Normal;}{\s1\keepn\widctlpar\adjustright \b\fs20\cgrid \sbasedon0 \snext0 heading 1;}{\s2\keepn\widctlpar\adjustright \b\fs18\cgrid " & _
        "\sbasedon0 \snext0 heading 2;}{\*\cs10 \additive Default Paragraph Font;}}{\info{\title Patient:}{\author webb}{\operator webb}{\creatim\yr2000\mo8\dy24\hr9\min33}{\revtim\yr2000\mo8\dy25\hr8\min46}{\version9}{\edmins12}{\nofpages1}{\nofwords25}" & _
        "{\nofchars145}{\*\company usgov}{\nofcharsws0}{\vern89}}\margl432\margr432\margt432\margb432 \widowctrl\ftnbj\aenddoc\hyphcaps0\formshade\viewkind1\viewscale100\pgbrdrhead\pgbrdrfoot \fet0\sectd \psz1\linex0\endnhere\sectdefaultcl {\*\pnseclvl1" & _
        "\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5" & _
        "\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang" & _
        "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}"
        
        
        If b2ndNameLine Then
            sFooter = sFooter & "\intbl\row\pard\trowd\trgaph72\trrh-220\trkeep \clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb \cellx11070\clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb \cellx1586\cellx8820\cellx18070\pard\widctlpar\intbl\pard\intbl\fs20\b " & "Name:" & sPatientName & " \cell\pard\intbl\cell \cell"
            sFooter = sFooter & "\intbl\row\pard\trowd\\trgaph36\trrh-220\trkeep \clvertalt\clbrdrt\brdrs\brdrw0 \cltxlrtb \cellx810\clvertalt\clbrdrt\brdrs\brdrw0 \cltxlrtb \cellx3420\clvertalt\clbrdrt" & _
                        "\brdrs\brdrw0 \cltxlrtb \cellx4050\clvertalt\clbrdrt\brdrs\brdrw0 \cltxlrtb \cellx6210\clvertalt\clbrdrt\brdrs\brdrw0 \cltxlrtb \cellx7380\clvertalt\clbrdrt\brdrs\brdrw0 \cltxlrtb " & _
                        "\cellx11070\pard\plain \widctlpar \adjustright \fs20\cgrid {\fs16 \cell }\pard\plain \s1\keepn\widctlpar\intbl\outlinelevel0\adjustright \b\fs20\cgrid {"
        Else
            sFooter = sFooter & "\intbl\row\pard\trowd\\trgaph36\trrh-220\trkeep \clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb \cellx810\clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb \cellx3420\clvertalt\clbrdrt" & _
                        "\brdrs\brdrw10 \cltxlrtb \cellx4050\clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb \cellx6210\clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb \cellx7380\clvertalt\clbrdrt\brdrs\brdrw10 \cltxlrtb " & _
                        "\cellx11070\pard\plain \widctlpar \adjustright \fs20\cgrid {\fs16 Name:\cell }\pard\plain \s1\keepn\widctlpar\intbl\outlinelevel0\adjustright \b\fs20\cgrid {" & sPatientName
        End If
        
        sFooter = sFooter & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 Sex:\cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright " & _
        "\b\fs18\cgrid {\fs16 " & sSex & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 Sponsor:\cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sSponsorName & "\cell }\pard\plain \widctlpar\intbl\adjustright " & _
        "\fs20\cgrid {\fs16 \row }\trowd \trgaph36\trrh-220\trkeep \clvertalt\cltxlrtb \cellx810\clvertalt\cltxlrtb \cellx3420\clvertalt\cltxlrtb \cellx4050\clvertalt\cltxlrtb \cellx6210\clvertalt\cltxlrtb \cellx7380\clvertalt\cltxlrtb \cellx11070\pard " & _
        "\widctlpar\intbl\adjustright {\fs16 FMP/SSN:\cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\fs20 " & sFMPSponsorSSN & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 Tel H:\cell " & sHomePhone & "}{\b\fs16 \cell " & _
        "}{\fs16 Rank:\cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sRank & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 \row }\trowd \trgaph36\trrh-200\trkeep \clvertalt\cltxlrtb \cellx810" & _
        "\clvertalt\cltxlrtb \cellx3420\clvertalt\cltxlrtb \cellx4050\clvertalt\cltxlrtb \cellx6210\clvertalt\cltxlrtb \cellx7380\clvertalt\cltxlrtb \cellx11070\pard \widctlpar\intbl\adjustright {\fs16 DOB:\cell }\pard\plain " & _
        "\s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\fs16 " & sDOB & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 Tel W:\cell " & sWorkPhone & "}{\b\fs16 \cell }{\fs16 Unit:\cell }\pard\plain " & _
        "\s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sUnit & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\b\fs16 \row }\pard \widctlpar\intbl\adjustright {\fs16 PCat:\cell }\pard\plain " & _
        "\s2\li-36\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sPCat & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 CS:\cell " & sCS & "\cell Outpt Rec. Rm:\cell }\pard\plain "
        
        sFooter = sFooter & _
        "\s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sOutptRecRm & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 \row }\pard \widctlpar\intbl\adjustright {\fs16 MC Status:\cell }\pard\plain " & _
        "\s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sMCStatus & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 WS:\cell " & sWS & "\cell PCM:\cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright " & _
        "\b\fs18\cgrid {\b0\fs16 " & sPCM & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\b\fs16 \row }\trowd \trgaph36\trrh-200\trkeep \clvertalt\clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx810\clvertalt\clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx3420\clvertalt" & _
        "\clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx4050\clvertalt\clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx6210\clvertalt\clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx7380\clvertalt\clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx11070\pard \widctlpar\intbl\adjustright {\fs16 Insurance:" & _
        "\cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sInsurance & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 \cell }{\b\fs16 \cell }{\fs16 Tel. PCM:\cell }\pard\plain " & _
        "\s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\b0\fs16 " & sTelPCM & "\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\fs16 \row }\trowd \trgaph144 \clvertalt\cltxlrtb \cellx11070\pard \widctlpar\intbl\adjustright {\b\fs6 \cell " & _
        "}\pard \widctlpar\intbl\adjustright {\fs12 \row }\trowd \trgaph72 \clvertalt\cltxlrtb \cellx8586\clvertalt\cltxlrtb \cellx8820\clvertalt\cltxlrtb \cellx11070\pard \qr\widctlpar\intbl\adjustright {\b\fs14 CHRONOLOGICAL RECORD OF MEDICAL CARE\cell }\pard " & _
        "\widctlpar\intbl\adjustright {\b\fs14 \cell }\pard\plain \s2\keepn\widctlpar\intbl\outlinelevel1\adjustright \b\fs18\cgrid {\fs14 STANDARD FORM 600 (REV. 5)" & _
        "\par }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid {\b\fs14 Prescribed by GSA and ICMR" & _
        "\par FIRMR (41 CFR) 201-45.505\cell }\pard \widctlpar\intbl\adjustright {\fs16 \row }\trowd \trgaph144 \clvertalt\cltxlrtb \cellx11070\pard \widctlpar\intbl\adjustright {\b\fs6 \cell }\pard \widctlpar\intbl\adjustright {\fs12 \row }\pard " & _
        "\qc\widctlpar\intbl\adjustright {\fs16 THIS INFORMATION IS PROTECTED BY THE PRIVACY ACT OF 1974 (PL-93-579). UNAUTHORIZED ACCESS " & _
        "\par TO THIS INFORMATION IS A VIOLATION OF FEDERAL LAW. VIOLATORS WILL BE PROSECUTED.}{\b\fs16 \cell }\pard \widctlpar\intbl\adjustright {\fs16 \row }\pard \qr\widctlpar\intbl\adjustright {\fs10 \cell }\pard \widctlpar\intbl\adjustright {\fs12 \row }\trowd " & _
        "\trgaph144 \clvertalt\cltxlrtb \cellx11070\pard \qr\widctlpar\intbl\adjustright {\fs16 Page " & sCurrentPage & " of " & sTotalPages & "\cell }\pard \widctlpar\intbl\adjustright {\fs16 \row }\pard \widctlpar\adjustright {" & _
        "\par }}"
    End If   'End mbGreenSwitch

    GetSF600Footer = sFooter

End Function

Public Function GetSF600Header(sTag1 As String, sTag2 As String, bFirstPage As Boolean) As String

    Dim sHeader As String
    
    sHeader = _
    "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f16\froman\fcharset238\fprq2 Times New Roman CE;}{\f17\froman\fcharset204\fprq2 Times New Roman Cyr;}" & _
    "{\f19\froman\fcharset161\fprq2 Times New Roman Greek;}{\f20\froman\fcharset162\fprq2 Times New Roman Tur;}{\f21\froman\fcharset186\fprq2 Times New Roman Baltic;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;" & _
    "\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;" & _
    "\red128\green128\blue128;\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\cgrid \snext0 Normal;}{\s1\keepn\widctlpar\adjustright \b\fs20\cgrid \sbasedon0 \snext0 heading 1;}{\s2\keepn\widctlpar\adjustright \b\fs18\cgrid " & _
    "\sbasedon0 \snext0 heading 2;}{\*\cs10 \additive Default Paragraph Font;}}{\info{\title Patient:}{\author webb}{\operator webb}{\creatim\yr2000\mo6\dy12\hr17\min50}{\revtim\yr2000\mo6\dy13\hr14\min25}{\version3}{\edmins1}{\nofpages1}{\nofwords25}" & _
    "{\nofchars145}{\*\company usgov}{\nofcharsws0}{\vern89}}\margl576\margr576\margt576\margb576 \widowctrl\ftnbj\aenddoc\hyphcaps0\formshade\viewkind1\viewscale100\pgbrdrhead\pgbrdrfoot \fet0\sectd \linex0\endnhere\sectdefaultcl {\*\pnseclvl1" & _
    "\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5" & _
    "\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang" & _
    "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\trowd \trgaph108\trleft-108\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \trbrdrh\brdrs\brdrw10 \trbrdrv" & _
    "\brdrs\brdrw10 \clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx2160\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 \cltxlrtb \cellx10710\pard\plain " & _
    "\s1\qc\keepn\widctlpar\intbl\outlinelevel0\adjustright \b\fs20\cgrid {HEALTH RECORD\cell CHRONOLOGICAL RECORD OF MEDICAL CARE\cell }\pard\plain \widctlpar\intbl\adjustright \fs20\cgrid  {\row } "
    
    If Not bFirstPage And Not (sTag1 = "" And sTag2 = "") Then
       sHeader = sHeader & "\b\fs16\cgrid " & sTag1 & " \cell " & sTag2 & " \cell }\pard\plain \widctlpar\intbl\adjustright \fs16\cgrid {\row } \pard \par \widctlpar\adjustright {"
    End If
    
    sHeader = sHeader & "\par }}"
    
    GetSF600Header = sHeader

End Function
Private Function IsMarkedYes(sText As String) As Boolean

    IsMarkedYes = False
    If sText <> "" Then
        If Mid(sText, 1, 1) = "Y" Then
            IsMarkedYes = True
        End If
    End If

End Function


Public Function GetCS(sUnitNumber As String, sContextNCID) As String

    Dim sStmt As String

    On Error GoTo ErrHandler
    GetCS = ""
    
   sStmt = "SELECT VALUE_TEXT AS CS " & _
           "FROM MMI_GENERICDATA " & _
           "WHERE TYPE_NCID = 204687 AND UNIT_NUMBER = " & sUnitNumber
    
    goDBSqlOps.Execute sStmt
    
    Do While Not goDBSqlOps.EOF
        If GetCS = "" Then
            GetCS = GetCS & goDBSqlOps.Value("CS")
        Else
            GetCS = GetCS & "," & goDBSqlOps.Value("CS")
        End If
        goDBSqlOps.MoveNext
    Loop

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "GetDOB.GetCS", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function


Public Function GetWS(sUnitNumber As String) As String

On Error GoTo ErrHandler
GetWS = ""

Dim sStmt As String
Dim sWorkStatus As String
Dim bDive As Boolean
Dim bFly As Boolean
Dim bJump As Boolean
Dim bMP As Boolean
Dim bSub As Boolean
Dim bMob As Boolean



    sStmt = "Select distinct mg.type_ncid, mg.value_text " & _
            "From " & _
                "mmi_genericdata mg " & _
            "Where " & _
                "mg.type_ncid in (" & _
                NCID_SWS_DIVING_STATUS & "," & _
                NCID_SWS_FLYING_STATUS & "," & _
                NCID_SWS_JUMPING_STATUS & "," & _
                NCID_SWS_MP & "," & _
                NCID_SWS_SUBMARINE & "," & _
                NCID_SWS_ON_MOBILITY & ") and " & _
                "mg.unit_number = " & sUnitNumber

    goDBSqlOps.Execute sStmt

    Do While Not goDBSqlOps.EOF
        If IsMarkedYes("" & goDBSqlOps.Value("value_text")) Then
            Select Case goDBSqlOps.Value("type_ncid")
                Case NCID_SWS_DIVING_STATUS
                    sWorkStatus = "Dive"
                Case NCID_SWS_FLYING_STATUS
                    sWorkStatus = "Fly"
                Case NCID_SWS_JUMPING_STATUS
                    sWorkStatus = "Jump"
                Case NCID_SWS_MP
                    sWorkStatus = "MP"
                Case NCID_SWS_SUBMARINE
                    sWorkStatus = "Sub"
                Case NCID_SWS_ON_MOBILITY
                    sWorkStatus = "Mob"
            End Select
        
            If GetWS = "" Then
                GetWS = GetWS & sWorkStatus
            Else
                GetWS = GetWS & "," & sWorkStatus
            End If
        End If
        
        goDBSqlOps.MoveNext
    Loop


GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "SF600Data.GetWS", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function



Public Sub GetPCM(sUnitNumber As String, sPCM As String, sTelPCM As String)

On Error GoTo ErrHandler
sPCM = ""
sTelPCM = ""

Dim sStmt As String

    If gobjshared.UseSQLServer Then
        sStmt = "select " & _
                "pr.name PCM, " & _
                "pr.duty_phone_1 Phone1, " & _
                "pr.duty_phone_2 Phone2 " & _
                "from " & _
                "provider pr, " & _
                "mmi_generic_id mgi " & _
                "where " & _
                "pr.pcm_code = mgi.id_value and " & _
                "mgi.id_type_ncid = 14501724 and " & _
                "mgi.unit_number = " & sUnitNumber
    Else
        sStmt = "select " & _
                "Trim(pr.name) PCM, " & _
                "Trim(pr.duty_phone_1) Phone1, " & _
                "Trim(pr.duty_phone_2) Phone2 " & _
                "from " & _
                "provider pr, " & _
                "mmi_generic_id mgi " & _
                "where " & _
                "pr.pcm_code = mgi.id_value and " & _
                "mgi.id_type_ncid = 14501724 and " & _
                "mgi.unit_number = " & sUnitNumber
    End If
    
    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF <> True Then
        sPCM = goDBSqlOps.Value("PCM")
        sTelPCM = goDBSqlOps.Value("Phone1")
        If goDBSqlOps.Value("Phone2") <> "" Then
            If sTelPCM <> "" Then sTelPCM = sTelPCM & ";"
            sTelPCM = sTelPCM & goDBSqlOps.Value("Phone2")
        End If
    End If

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "SF600Data.GetPCM", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Sub
End Sub


Public Function GetPCMGroup(sUnitNumber As String) As String

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetPCMGroup = ""
    Exit Function

    sStmt = "select Trim(pr.name) PCM " & _
            "from provider pr, mmi_genericdata mgd" & _
            "where pr.pcm_code = mgd.value_text and " & _
            "type_ncid = 14501724 and " & _
            "mgd.unit_number = " & sUnitNumber
    
    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF <> True Then
        GetPCMGroup = goDBSqlOps.Value("PCM")
    End If

    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "SF600Data.GetPCMGroup", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function


Public Function GetContextNCID(sFacilityNCID As String) As String

On Error GoTo ErrHandler
GetContextNCID = ""

Dim sStmt As String

    sStmt = "select context_ncid " & _
            "From facility_ncid_correlation " & _
            "Where " & _
                "type = 'ADT' and " & _
                "facility_ncid = " & sFacilityNCID

    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF <> True Then
        GetContextNCID = Trim(goDBSqlOps.Value("context_ncid"))
    End If
    

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "SF600Data.GetContextNCID", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function

Public Function BuildSF600(sPatientUnitNumber As String, _
                           sFacilityNCID As String, _
                           sRTF As String, _
                           bPreview As Boolean, _
                           bSkipPrinterDialog As Boolean, _
                           Optional oEnc As Object = Nothing) As Boolean

    On Error GoTo ErrHandler
    BuildSF600 = False
    
    Dim i As Integer
    Dim iNumberOfCopies As Integer
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
    
    Dim sHeaderTag1 As String
    Dim sHeaderTag2 As String
    
    Dim xxx As String
    Dim sTemp As String
    Dim sStatus As String
    Dim sRelToPat As String
    Dim nCurrentEnctrStatus As Integer
    Dim sMedicalRecordNumber As String
    Dim sHeader As String
    Dim sHeader2 As String
    Dim sContextNCID As String
    Dim sSponsorUnitNumber As String
    
    Dim iMice As Integer
    iMice = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    #If coiDebug Then
         Dim DR As DLog.DebugRec
         gCOI_DL.DebugLog "COI-MOP#2-1-1", "", "", "Begin", DR
    #End If

    '-- Get Data
    '----------------------------------------------------------------------------
    '-- Patient Demographics
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, moPatientRec) Then GoTo CleanUp

    '-- Sponsor Demographics
    '----------------------------------------------------------
    If gobjshared.IsAppMode(modeCHCSI_GUI) Then
        sSponsorUnitNumber = moPatientRec.Fields.Item("Sponsor_UNIT_NUMBER").Value
    Else
        sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(moPatientRec.Fields.Item("Sponsor_SSN").Value & "")
    End If
    
    If sSponsorUnitNumber = sPatientUnitNumber Then
       Set moSponsorRec = moPatientRec.Clone
    Else
       If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, moSponsorRec) Then GoTo CleanUp
    End If
    
    DoEvents

    '-- Context NCID
    If Not gobjshared.IsAppMode(modeCHCSI_GUI) Then
        gsContextNCID = GetContextNCID(sFacilityNCID)
    End If
    
    '-- SF600 HEADER
    '------------------------------------------------------------------------------
    sHeaderTag1 = ""
    sHeaderTag2 = ""
    If Not oEnc Is Nothing Then
       sHeaderTag1 = Format(oEnc.startdts, gsDefaultDateTimeFormat)
       sHeaderTag2 = "Facility: " & oEnc.Facility & vbTab & "Clinic: " & oEnc.Clinic & vbTab & " Provider: " & oEnc.Providers(1).FullName
    End If
    
    sHeader = GetSF600Header(sHeaderTag1, sHeaderTag2, True)
    sHeader2 = GetSF600Header(sHeaderTag1, sHeaderTag2, False)
    
    frmSF600.txHeader.PageWidth = 0 '- reset fir autosize
    frmSF600.txHeader.PageHeight = 0 '- reset fir autosize
    frmSF600.txHeader.Text = ""
    frmSF600.txHeader.Height = 1400
    frmSF600.txHeader.PageMarginT = 450
    frmSF600.txHeader.PageMarginL = 600

    '--- SCR 17046
    sBody = Replace(sRTF, "\b\qc", "\b\ql")
    
    DoEvents

'-- SF600 BODY
'------------------------------------------------------------------------------
    frmSF600.txBody.PageWidth = 0 '- reset fir autosize
    frmSF600.txBody.PageHeight = 0 '- reset fir autosize
    frmSF600.txBody.AutoExpand = True
    frmSF600.txBody.Text = ""
    frmSF600.txBody.Height = 500
    
    frmSF600.txBody.PageMarginT = 1250
    frmSF600.txBody.PageMarginL = 1440
  If Me.mbGreenSwitch Then  'hwcc 2/11/2003 - this is throw away switch when GREEN switch no longer valid
    frmSF600.txBody.PageMarginB = 2500    'smaller Footer
  Else
    frmSF600.txBody.PageMarginB = 3220
  End If
    frmSF600.txBody.PageHeight = 15748
    frmSF600.txBody.PageWidth = 12240

    sUpdatedBody = FixImageScale(sBody, 15748)
    frmSF600.txBody.RTFSelText = sUpdatedBody
    
    DoEvents
    
'-- SF600 FOOTER
'------------------------------------------------------------------------------
    frmSF600.txFooter.PageWidth = 0 '- reset fir autosize
    frmSF600.txFooter.PageHeight = 0 '- reset fir autosize
    frmSF600.txFooter.Text = ""
    frmSF600.txFooter.AutoExpand = True
    frmSF600.txFooter.RTFSelText = sFooter
  If Me.mbGreenSwitch Then  'hwcc 2/11/2003 - this is throw away switch when GREEN switch no longer valid
    frmSF600.txFooter.PageMarginT = 13248
  Else
    frmSF600.txFooter.PageMarginT = 12550
  End If
    frmSF600.txFooter.PageMarginL = 600

    iPages = frmSF600.txBody.CurrentPages
    
    DoEvents

'-- PRINT PREVIEW
'-------------------------------------------------------------------------------
    If gobjshared.UseSQLServer And Not gobjshared.IsAppMode(modeCHCSI_GUI) Then
        frmSF600.txWaterMark.ObjectInsertAsChar 0, gobjshared.ImagesPath & "\TRAINING.TIF", -1, 250, 250, &H2&, 1
        DoEvents
    End If
    If bPreview Then

        '-- Give Form information to print the SF600 by calling this procedure
        Set frmPreviewSF600.oCalledBy = Me
        frmPreviewSF600.sFacilityNCID = sFacilityNCID
        frmPreviewSF600.sUnitNumber = sPatientUnitNumber
        frmPreviewSF600.sRTF = sRTF
        Set frmPreviewSF600.oEnc = oEnc

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
            
            If i = 0 Then
                frmSF600.txHeader.RTFSelText = sHeader
            Else
                frmSF600.txHeader.RTFSelText = sHeader2
            End If
            
            frmSF600.txHeader.ZOrder 1
            frmSF600.txHeader.PrintDevice = frmPreviewSF600.PicPreview(i).hDC
            frmSF600.txHeader.PrintPage 1

            frmSF600.txBody.PrintDevice = frmPreviewSF600.PicPreview(i).hDC
            frmSF600.txBody.PrintPage i + 1

            frmSF600.txFooter.ZOrder 1
            frmSF600.txFooter.Text = ""
            frmSF600.txFooter.RTFSelText = GetSF600Footer(sPatientUnitNumber, CStr(i + 1), CStr(iPages))
            frmSF600.txFooter.PrintDevice = frmPreviewSF600.PicPreview(i).hDC
            frmSF600.txFooter.PrintPage 1

        Next i

        #If coiDebug Then
             gCOI_DL.DebugLog "COI-MOP#2-1-1", "", "", "End", DR
        #End If

        Screen.MousePointer = iMice
        
        frmPreviewSF600.Show vbModal

'-- PRINT
'-------------------------------------------------------------------------------
    Else
        If Printers.Count = 0 Then
           MsgBxARMd "No printers are installed. To install a printer, point to Settings on the Windows Start menu, " & _
                "click Printers, and then double-click Add Printer. Follow the instructions in the wizard.", vbOKOnly + vbExclamation
           Exit Function
        End If
        
        'Set default copies and page range for Printer
        iNumberOfCopies = 1
        lPageFrom = 0
        lPageTo = iPages - 1
        
        If bSkipPrinterDialog Then GoTo PrintDirect  'User choose not to show Printer Dialog Window
        
        frmPrint.MaxPageTo = iPages
        
        Screen.MousePointer = vbDefault
        frmPrint.Show vbModal
        Screen.MousePointer = vbHourglass
        
        If frmPrint.PrintOperationCancelled Then
           Unload frmPrint
           GoTo CleanUp
        End If
        iNumberOfCopies = frmPrint.NumberOfCopies
        If Not frmPrint.PrintAll Then
           lPageFrom = frmPrint.PageFrom - 1
           lPageTo = frmPrint.PageTo - 1
        End If
        Unload frmPrint
        
PrintDirect:
        For iPrintCount = 1 To iNumberOfCopies
            For i = lPageFrom To lPageTo
                Printer.Print
                
                If gobjshared.UseSQLServer Then
                    frmSF600.txWaterMark.PrintDevice = Printer.hDC
                    frmSF600.txWaterMark.PrintPage 1
                End If
    
                '--
                frmSF600.txHeader.Text = ""
                If i = lPageFrom Then
                    frmSF600.txHeader.RTFSelText = sHeader
                Else
                    frmSF600.txHeader.RTFSelText = sHeader2
                End If
                frmSF600.txHeader.ZOrder 1
                frmSF600.txHeader.PrintDevice = Printer.hDC
                frmSF600.txHeader.PrintPage 1
    
                '-- Print body
                frmSF600.txBody.PrintDevice = Printer.hDC
                frmSF600.txBody.PrintPage i + 1
    
                '-- print footer
                frmSF600.txFooter.Text = ""
                frmSF600.txFooter.RTFSelText = GetSF600Footer(sPatientUnitNumber, CStr(i + 1), CStr(iPages))
                frmSF600.txFooter.ZOrder 1
                frmSF600.txFooter.PrintDevice = Printer.hDC
                frmSF600.txFooter.PrintPage 1
                            
                Printer.NewPage
            Next i
        Next iPrintCount


        Printer.EndDoc
        BuildSF600 = True

        #If coiDebug Then
             gCOI_DL.DebugLog "COI-MOP#2-1-1", "", "", "End", DR
        #End If

        GoTo CleanUp
    End If     '--  If bPreview Then

    If PrintPreviewCancel = "Cancel" Then BuildSF600 = True
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
     gobjshared.ShowVBError Err.Number, Err.Description, "FormSF600.BuildSF600", "CHCSII_Forms", vbExclamation
  End If
CleanUp:
    Screen.MousePointer = iMice
    Unload frmPreviewSF600
    Set frmPreviewSF600 = Nothing
    Unload frmSF600
    Set frmSF600 = Nothing
End Function

Private Sub Class_Initialize()
   Set moPatientDemographics = New PatientDemographics
End Sub

Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
    Set moPatientRec = Nothing
    Set moSponsorRec = Nothing     'hwcc 2/19/2003 - SCR 32681, someone forget to release this reference
End Sub


