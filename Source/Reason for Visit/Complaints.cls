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
Attribute VB_Name = "Complaints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Complaint"
Attribute VB_Ext_KEY = "Member0" ,"Complaint"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable to hold collection
Private mCol As Collection

'this is a completely private copy of the collection. This is used so that we can know what items
'have been marked as deleted from the collection.  This is necessary since the item will need to
'be deleted from the Complaints object by hand.
Private mCol2 As Collection
Private msEncounterId As String
Private moProblem As MMMHISProblems.Problem
Private moProblemOps As MMMHISProblems.ProblemOps
Private moOneProblem As OneProblem
Private moComplaints As MMMHISProblems.IComplaints
Private msAppointmentId As String

'This is necessary for the Triage Application so that the BI module can tell me what appt it is using
Public Property Let AppointmentId(vData As String)
    msAppointmentId = vData
End Property

'--------------------------------------------------------------------------------------------------------------------------
'Returns complaints for a patient.  Pass in the EncounterId and receive all complaints for that Encoutner
'--------------------------------------------------------------------------------------------------------------------------
Public Function Load(EncoutnerId As Long) As Boolean
    Dim x As MMMHISProblems.IComplaints
    Dim oComp As Complaint
    Dim i As Integer
    Dim cArray As Variant
On Error GoTo errorHandler
    
    Set mCol = New Collection
    Set mCol2 = New Collection
    
    msEncounterId = EncoutnerId
    
    Set x = moComplaints.Filter(EncoutnerId, NCID_ACTIVE)
    
    For i = 1 To x.Count
        Set oComp = New Complaint

        If x.Item(i).Comment <> "" Then
            If x.Item(i).SnoID = "-1" Then
                oComp.Comment = x.Item(i).Comment
            Else
                cArray = Split(x.Item(i).Comment, "|", , vbTextCompare)
                oComp.Comment = cArray(0) & "|" & cArray(1)
                oComp.BIIndex = cArray(2)
            End If
        End If
        oComp.ComplaintText = x.Item(i).Complaint
        oComp.EncounterFacilityNCID = x.Item(i).EncounterFacilityNCID
        oComp.EncounterNumber = x.Item(i).EncounterNumber
        oComp.SnoID = x.Item(i).SnoID
        oComp.PrefixCode = x.Item(i).prefix
        If x.Item(i).NewFollowUp = "New" Then
            oComp.visittype = enVisitTypeNew
        Else
            oComp.visittype = enVisitTypeFollowup
        End If
        'oComp.StatusNCID = x.Item(i).StatusNCID
        oComp.DataId = x.Item(i).DataId
        mCol.Add oComp
        mCol2.Add oComp
    Next
    
    Set oComp = Nothing
    
    Load = True
    Exit Function
errorHandler:
    If Err.number = 9 Then Resume Next
    goBB.SharedObj.ShowVBError Err.number, Err.Description, "GetComplaints", "ReasonForVisit.Complaints", vbCritical
    Resume Next
    Load = False
End Function

Public Function SaveComplaints() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim oComp As Complaint
    Dim x As MMMHISProblems.IComplaints
    Dim sComplaints As String
    Dim cArray As Variant
    Dim sComments As String
On Error GoTo errorHandler
    
    Set x = moComplaints.Filter(msEncounterId, NCID_ACTIVE)
    
    For i = 1 To mCol2.Count
        Set oComp = mCol2.Item(i)
        
        If oComp.StatusNCID = NCID_DELETED_ERROR Then 'delete
            For j = 1 To x.Count
                If oComp.DataId = "" Then Exit For  'this is an item that was entered and deleted in the same session
                If x.Item(j).DataId = oComp.DataId Then
                    x.Item(j).StatusNCID = NCID_DELETED_ERROR
                    x.Item(j).Update
                    Exit For
                End If
            Next
        ElseIf oComp.StatusNCID = NCID_MODIFIED Then    'update
            For j = 1 To x.Count
                If x.Item(j).DataId = oComp.DataId Then
                    x.Item(j).NewFollowUp = oComp.VisitTypeText
                    x.Item(j).Comment = oComp.Comment & "|" & oComp.BIIndex
                    x.Item(j).Update
                    Exit For
                End If
            Next
        ElseIf oComp.StatusNCID = NCID_ACTIVE Then 'perform addnew
'            Set oComp = mCol2.Item(i)
            
'            If goBB.SharedObj.CmdLineSwitches.Item("GEMS_ARMY") <> "" Then
'                If oComp.Comment <> "" Then
'                    cArray = Split(oComp.Comment, "|", , vbTextCompare)
'                    sComments = cArray(1)
'                End If
'            End If
            
            With oComp
                ' SCR #28661;   Developer: Jane Shen 02/25/2003 05:55 PM
                If .SnoID = "-1" Then
                    moComplaints.AddNew CLng(mcsComplaintNCID), .SnoID, .PrefixCode, .Comment & .BIIndex, , .VisitTypeText, .EncounterFacilityNCID, , .EncounterFacilityNCID, Now, .EncounterNumber
                Else
                    moComplaints.AddNew CLng(mcsComplaintNCID), .SnoID, .PrefixCode, .Comment & "|" & .BIIndex, , .VisitTypeText, .EncounterFacilityNCID, , .EncounterFacilityNCID, Now, .EncounterNumber
                End If
            End With
        ElseIf oComp.SnoID = "-1" Then   'comments only
            For j = 1 To x.Count
                If x.Item(j).SnoID = -1 Then
                    x.Item(j).Comment = oComp.Comment
                    x.Item(j).Update
                    Exit For
                End If
            Next
        End If
        
    Next

    For i = 1 To mCol.Count
        If i > 1 And Len(sComplaints) > 0 Then
            sComplaints = sComplaints & ", " & mCol.Item(i).ComplaintText
        Else
            sComplaints = sComplaints & mCol.Item(i).ComplaintText
        End If
    Next
    
    'SCR 24097: M.Broaddus 6/10/2002: Need to send the selected RFV's to the appts module so that they can be shown there.
    ' SCR #50577;   Developer: MAG 03/09/2004
    If Not goBB.CurrentEncounter Is Nothing Then 'And msAppointmentId = "" Then
        'If Not goBB.CurrentEncounter Is Nothing Then msAppointmentId = goBB.CurrentEncounter.AppointmentId
        msAppointmentId = goBB.CurrentEncounter.AppointmentId
        If goBB.SharedObj.AppMode = modeTheater And msAppointmentId <> "" Then
            goBB.Comm.Message cwiGENERIC, "UPDATE_APPT_RFV|" & msAppointmentId & "|" & sComplaints, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_REASON_FOR_VISIT
        End If
    End If
    Exit Function
    
errorHandler:
    goBB.SharedObj.ShowVBError Err.number, Err.Description, "SaveComplaints", "ReasonForVisit.Complaints", vbCritical
End Function

Public Function GetRTF(EncounterId As Long) As String
    
    Dim rtfString As String
    Dim i As Long
    Dim sTemp As String
    Dim sComplaint As String
    Dim sTimeZone As String
    Dim sSQL As String
    Dim sRTFNote As String
    Dim objComplaint As Complaint
    Dim sComments As String
    Dim msRFVComments As String
    Dim FTZ As String   'full time zone
    Dim ATZ As String   'abvreiated time zone
    Dim sListBox As String
    Dim sSummaryComment As String 'This is the complaint comment that is entered into the seperate text popup
    Dim cArray As Variant   'array to split comment string into
    Dim apptComments As String
    
On Error GoTo errorHandler
    
    rtfString = ""

    If goBB.SharedObj.CmdLineSwitches.Item("GEMS_ARMY") <> "" Then
        sComplaint = GetChiefComplaintApptComment(EncounterId, apptComments)
    Else
        sComplaint = Trim$(GetChiefComplaint)
    End If
    
    gDataDate = Now
    goBB.Comm.GetLocalTimeZone sTimeZone, ATZ
    
    Load EncounterId
    
    If goBB.Comm.CmdLineSwitch("GREEN") = "GREEN" Then
    
    ' SCR #28661;   Developer: Jane Shen 02/25/2003 05:56 PM
    '--------------------------------  For GREEN ----------------------
    If Len(sComplaint) <> 0 Then
        'Escape out RTF Syntax from input string
        sComplaint = goBB.SharedObj.RTFCodeFilter(sComplaint)
        
        ' Begin:SCR #38059;   Developer: Jane Shen 06/23/2003 05:08 PM
        ' Note: Function Changed:    change made in conplaints.cls:GetRTF()    Impact of Change:     Put space before the reason for Appointment, other wise font will get very big with leading numbers.
        ' End: SCR #38059;
        If goBB.EncParent Is Nothing Then
            ' SCR #45654;   Developer:  12/05/2003
            'Removed "\ul"
            rtfString = "\plain\f2\fs16\b Reason For Appointment: \plain\f2\fs16 " & sComplaint & " \par "
        Else
            If Not goBB.CurrentEncounter.EncounterType = TeleConsult Then
                rtfString = "\plain\f2\fs16\b Reason For Appointment: \plain\f2\fs16 " & sComplaint & " \par "
            Else
                rtfString = "\plain\f2\fs16\b Reason For Appointment Call: \plain\f2\fs16 " & sComplaint & " \par "
            End If
        End If
    End If
    
    ' Begin:SCR #38921;   Developer: Jane Shen 07/25/2003 02:59 PM
    ' Note: Function Changed:    change made in Complaints.cls:GetRTF()    Impact of Change:     With or without Resaon for Visit, Allery is get verified.
    ' End: SCR #38921;
    If goBB.CurrentEncounter Is Nothing Then
    'nothing
    Else
        If goBB.CurrentEncounter.AlleryVerifiedByDATE = 0 Then
            goBB.CurrentEncounter.AlleryVerifiedByDATE = gDataDate
        End If
        
        If Len(goBB.CurrentEncounter.AlleryVerifiedByNCID) And goBB.CurrentEncounter.AlleryVerifiedByNCID <> "0" Then
            goBB.GEMS_Concept.UniqueId = goBB.CurrentEncounter.AlleryVerifiedByNCID
            sTemp = Trim(goBB.GEMS_Concept.PrefRep("2000").Representation)
            
            If rtfString = "" Then
                rtfString = "\plain\f2\fs16\b\ul"
            End If
            ' Begin:SCR #42251;   Developer: Jane Shen 09/24/2003 03:10 PM
            ' Note: Function Changed:     change made in Complaints.cls:GetRTF()     Impact of Change:      There are two issues in the this tracker.       1. Fixed the problem with the wrong time Zone display, should be EDT, CDT.       2. The verify allergy problem was fixed with SCR #38062 which is in build 153.
            ' End: SCR #42251;
            If Len(Trim$(sTemp)) = 0 Then
                rtfString = rtfString & "\par Allergen information verified @ " & FormatDTG(CStr(goBB.CurrentEncounter.AlleryVerifiedByDATE)) & " " & ATZ & " \par "
            Else
                'Escape out RTF Syntax from input string
                sTemp = goBB.SharedObj.RTFCodeFilter(sTemp)
                rtfString = rtfString & "\par Allergen information verified by " & sTemp & " @ " & FormatDTG(CStr(goBB.CurrentEncounter.AlleryVerifiedByDATE)) & " " & ATZ & " \par "
            End If
        End If
    End If
    
    If mCol.Count > 0 Then
        ' Begin:SCR #38062;   Developer: Jane Shen 07/01/2003 02:26 PM
        ' Note: Function Changed:    changes made in Complains.cls:GetRTF()    Impact of Change:     Removed "\par"
        ' End: SCR #38062;
        rtfString = rtfString & "\plain\f2\fs16\b Reason(s) For Visit (Chief Complaint): "
    End If
    
    For i = 1 To mCol.Count
        Set objComplaint = mCol.Item(i)
        
        If objComplaint.SnoID <> "-1" Then
            If InStr(1, objComplaint.Comment, "|", 1) > 0 Then
                If objComplaint.Comment <> "" Then
                    cArray = Split(objComplaint.Comment, "|", , vbTextCompare)
                    'sListBox = Left(objComplaint.Comment, InStr(1, objComplaint.Comment, "|", 1) - 1)
                    'sComments = Right(objComplaint.Comment, Len(objComplaint.Comment) - InStr(1, objComplaint.Comment, "|", 1))
                    sComments = cArray(1)
                End If
            Else
               sListBox = 0
               sComments = objComplaint.Comment
            End If
            
            'Business Rule: If we have a BIIndex for the complaint and the complaint text is "" then we need to use
            'the comment as the complaint text.  This happens only if the SnoID sent to us from the body image
            'module is not a problem but a procedure, as in the case of a tournaquet (sp?).
            'Remember: Body Image complaints do not have comments of their own.
            
            'Army Triage -
            If sComments <> "" Then
                'Escape out RTF Syntax from input string
                sComments = goBB.SharedObj.RTFCodeFilter(": " & sComments)
            End If
            If goBB.SharedObj.CmdLineSwitches.Item("GEMS_ARMY") <> "" Then
                If objComplaint.BIIndex <> "" And objComplaint.ComplaintText = "" Then
                    rtfString = rtfString & "\plain\f2\fs16 " & sComments & _
                        " (" & goBB.SharedObj.RTFCodeFilter(objComplaint.VisitTypeText) & _
                        ")  ; "
                Else
                    If sComments = objComplaint.ComplaintText Then sComments = ""
                    rtfString = rtfString & "\plain\f2\fs16 " & goBB.SharedObj.RTFCodeFilter(objComplaint.ComplaintText) & _
                        " (" & goBB.SharedObj.RTFCodeFilter(objComplaint.VisitTypeText) & _
                        ") " & sComments & "; "
                End If
            Else
                rtfString = rtfString & "\plain\f2\fs16 " & goBB.SharedObj.RTFCodeFilter(objComplaint.ComplaintText) & _
                        " (" & goBB.SharedObj.RTFCodeFilter(objComplaint.VisitTypeText) & _
                        ") " & sComments & "; "
            End If
                    
        Else
            sSummaryComment = goBB.SharedObj.RTFCodeFilter(objComplaint.Comment)  'This is an overall comment that was entered into the popup window
        End If
                    
    Next
    
   
    If Not Len(Trim$(sSummaryComment)) = 0 Then
       rtfString = rtfString & "\par\plain\f2\fs16\b Comments: \plain\f2\fs16  " & sSummaryComment & " \par"
    Else
        rtfString = rtfString & " \par"
    End If
    
    If goBB.SharedObj.CmdLineSwitches.Item("GEMS_ARMY") <> "" And apptComments <> "" Then
        rtfString = rtfString & "\par\b Appointment Note \par\b0 " & goBB.SharedObj.RTFCodeFilter(apptComments) & " \par"
    End If
   
    If Not Len(Trim$(rtfString)) = 0 Then
       sRTFNote = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}" _
       & "{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fprq2 Arial;}}" _
       & "{\colortbl\red0\green0\blue0;}" _
       & "\deflang1033\horzdoc{\*\fchars }{\*\lchars }" _
       & "\pard\tx3690\tx6480\plain\f2\fs16\b\ul" _
       & "\pard\plain\f2\fs16 " _
       & rtfString & "\pard\tx3690\tx6480\pard }"
' SCR #34076;   Developer:  03/27/2003 11:14 AM
' Note: Function Changed:    Change made in Complaints.cls:GetRTF()    Impact of Change:     change "Par" to "Pard"

       GetRTF = sRTFNote
    End If

    Else
    '----------------------------  END OF GREEN ----------------------------------
    If Len(sComplaint) <> 0 Then
        If goBB.EncParent Is Nothing Then
            rtfString = "\plain\f2\fs16\b\ul Appointment Reason For Visit: " & sComplaint & " \par "
        Else
            If Not goBB.CurrentEncounter.EncounterType = TeleConsult Then
                rtfString = "\plain\f2\fs16\b\ul Appointment Reason For Visit: " & sComplaint & " \par "
            Else
                rtfString = "\plain\f2\fs16\b\ul Appointment Reason For Call: " & sComplaint & " \par "
            End If
        End If
    End If
    
    If goBB.CurrentEncounter Is Nothing Then
    'nothing
    Else
        If goBB.CurrentEncounter.AlleryVerifiedByDATE = 0 Then
            goBB.CurrentEncounter.AlleryVerifiedByDATE = gDataDate
        End If
    
        If Len(goBB.CurrentEncounter.AlleryVerifiedByNCID) And goBB.CurrentEncounter.AlleryVerifiedByNCID <> "0" Then
            goBB.GEMS_Concept.UniqueId = goBB.CurrentEncounter.AlleryVerifiedByNCID
            sTemp = Trim(goBB.GEMS_Concept.PrefRep("2000").Representation)
            
            If rtfString = "" Then
                rtfString = "\plain\f2\fs16\b\ul"
            End If
            If Len(Trim$(sTemp)) = 0 Then
                rtfString = rtfString & "\par Allergen information verified @ " & FormatDTG(CStr(goBB.CurrentEncounter.AlleryVerifiedByDATE)) & " " & ATZ & " \par "
            Else
                sTemp = goBB.SharedObj.RTFCodeFilter(sTemp)
                rtfString = rtfString & "\par Allergen information verified by " & sTemp & " @ " & FormatDTG(CStr(goBB.CurrentEncounter.AlleryVerifiedByDATE)) & " " & ATZ & " \par "
            End If
        End If
    End If
    
    If mCol.Count > 0 Then
        rtfString = rtfString & "\par Selected Reason(s) For Visit: \par "
    End If

    For i = 1 To mCol.Count
        Set objComplaint = mCol.Item(i)

        If objComplaint.SnoID <> "-1" Then
            If InStr(1, objComplaint.Comment, "|", 1) > 0 Then
                If objComplaint.Comment <> "" Then
                    cArray = Split(objComplaint.Comment, "|", , vbTextCompare)
                    'sListBox = Left(objComplaint.Comment, InStr(1, objComplaint.Comment, "|", 1) - 1)
                    'sComments = Right(objComplaint.Comment, Len(objComplaint.Comment) - InStr(1, objComplaint.Comment, "|", 1))
                    sComments = cArray(1)
                End If
            Else
               sListBox = 0
               sComments = objComplaint.Comment
            End If

            'Business Rule: If we have a BIIndex for the complaint and the complaint text is "" then we need to use
            'the comment as the complaint text.  This happens only if the SnoID sent to us from the body image
            'module is not a problem but a procedure, as in the case of a tournaquet (sp?).
            'Remember: Body Image complaints do not have comments of their own.
            
            sComments = goBB.SharedObj.RTFCodeFilter(sComments)
            'Army Triage -
            If goBB.SharedObj.CmdLineSwitches.Item("GEMS_ARMY") <> "" Then
                If objComplaint.BIIndex <> "" And objComplaint.ComplaintText = "" Then
                    rtfString = rtfString & "\plain\f2\fs16 " & sComments & _
                        " (" & goBB.SharedObj.RTFCodeFilter(objComplaint.VisitTypeText) & _
                        ") Comments: \par"
                Else
                    If sComments = objComplaint.ComplaintText Then sComments = ""
                    rtfString = rtfString & "\plain\f2\fs16 " & objComplaint.ComplaintText & _
                        " (" & goBB.SharedObj.RTFCodeFilter(objComplaint.VisitTypeText) & _
                        ") Comments: " & sComments & "\par"
                End If
            Else
                rtfString = rtfString & "\plain\f2\fs16 " & goBB.SharedObj.RTFCodeFilter(objComplaint.ComplaintText) & _
                        " (" & goBB.SharedObj.RTFCodeFilter(objComplaint.VisitTypeText) & _
                        ") Comments: " & sComments & "\par"
            End If

        Else
            sSummaryComment = goBB.SharedObj.RTFCodeFilter(objComplaint.Comment)  'This is an overall comment that was entered into the popup window
        End If

    Next

    If Not Len(Trim$(sSummaryComment)) = 0 Then
       rtfString = rtfString & "\par \plain\f2\fs16 Comments: " & sSummaryComment & " \par"
    End If
    
    If goBB.SharedObj.CmdLineSwitches.Item("GEMS_ARMY") <> "" And apptComments <> "" Then
        rtfString = rtfString & "\par\b Appointment Note \par\b0 " & goBB.SharedObj.RTFCodeFilter(apptComments) & " \par"
    End If

    If Not Len(Trim$(rtfString)) = 0 Then
       sRTFNote = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}" _
       & "{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fprq2 Arial;}}" _
       & "{\colortbl\red0\green0\blue0;}" _
       & "\deflang1033\horzdoc{\*\fchars }{\*\lchars }" _
       & "\pard\tx3690\tx6480\plain\f2\fs16\b\ul" _
       & "\par \pard\plain\f2\fs16 " _
       & rtfString & "\pard\tx3690\tx6480\pard }"
       GetRTF = sRTFNote
    End If

    End If  '--------------------- NOT GREEN --------------

    Exit Function
errorHandler:
    goBB.SharedObj.ShowVBError Err.number, Err.Description, "GetRTF", "ReasonForVisit.Complaints", vbCritical
End Function

Public Function Add(PrefixCode As String, Comment As String, visittype As enVisitType, SnoID As String, Complaint As String, EncounterFacilityNCID As String, EncounterNumber As String, BIIndex As String, Optional sKey As String) As Complaint
    Dim objNewMember As Complaint
    Set objNewMember = New Complaint
On Error GoTo errorHandler

    'set the properties passed into the method
    objNewMember.PrefixCode = PrefixCode
    objNewMember.Comment = Comment
    objNewMember.visittype = visittype
    objNewMember.SnoID = SnoID
    objNewMember.ComplaintText = Complaint
    objNewMember.EncounterFacilityNCID = EncounterFacilityNCID
    objNewMember.EncounterNumber = EncounterNumber
    objNewMember.StatusNCID = NCID_ACTIVE
    objNewMember.BIIndex = BIIndex
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
        mCol2.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
        mCol2.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    Exit Function
errorHandler:
    goBB.SharedObj.ShowVBError Err.number, Err.Description, "Add", "ReasonForVisit.Complaints", vbCritical
End Function

Public Property Get Item(vntIndexKey As Variant) As Complaint
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    Dim objRmvMember As Complaint
    Dim objComp As Complaint
    Dim i As Integer
    ' SCR #34637;   Developer: Jane Shen 04/17/2003 02:39 PM
    ' Note: Function Changed:    change made in frmReasonForVisit.frm:cmdRemove_Click(), and  Complaints.cls:Remove()    Impact of Change:    have to match with the SnoID.
    
    Set objRmvMember = Item(vntIndexKey)
    For i = 1 To mCol2.Count
        Set objComp = mCol2(i)
        If objRmvMember.SnoID = objComp.SnoID Then  'find the match SnoID
            If objComp.DataId <> "" And objComp.StatusNCID <> NCID_DELETED_ERROR Then
                objComp.StatusNCID = NCID_DELETED_ERROR
                Exit For
            ElseIf objComp.DataId = "" Then
                mCol2.Remove i
                Exit For
            End If
        End If
    Next
    
    For i = 1 To mCol.Count
        Set objComp = mCol(i)
        If objRmvMember.SnoID = objComp.SnoID Then  'find the match SnoID
            mCol.Remove i
            Exit For
        End If
    Next
    
'    mCol.Remove vntIndexKey
'    mCol2.Item(vntIndexKey).StatusNCID = NCID_DELETED_ERROR  'indicate on the local copy that this item needs to be taken out

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mCol2 = New Collection
    Set moProblem = goBB.Comm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
    Set moComplaints = moProblem.ProblemOps.Complaints
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mCol2 = Nothing
End Sub

