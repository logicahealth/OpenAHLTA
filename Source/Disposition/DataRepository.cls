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
Attribute VB_Name = "DataRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CodedValues"
Attribute VB_Ext_KEY = "Member1" ,"MetaData"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"RTFs"
Option Explicit
'maint: WRogers 1/09/2003
' Added SetClinicCapabilities(), MEPRSCode, DMIScode and AmbulatoryAuthorization
'

Private mvarCodedValues             As CodedValues
Private mvarMetaData                As MetaData
Private mvarRTFs                    As RTFs
Private mbAmbulatoryAuthorization   As Boolean
Private msMEPRSCode                 As String
Private msDMISIdCode                As String
'Private mobjMMI                     As MmiDoc
'Private mobjMember                  As Member
Private mbUpdated                   As Boolean

'Convert a code for the Legacy system to a NCID.
Private Function DispLegacytoNCID(LegacyCode As String) As String

    Select Case LegacyCode
        Case "95"                  '"204849"
            DispLegacytoNCID = NCID_Admitted
            
        Case "96"                   '"549"
            DispLegacytoNCID = NCID_Expired
            
        Case "90"          '"204848"
            DispLegacytoNCID = NCID_ImmediateReferal
        
        Case "06"        '"546"
            DispLegacytoNCID = NCID_LeftAgainstAdvice
        
        Case "97" '"13"
            DispLegacytoNCID = PatientLeftWithoutBeingSeen
        
        Case "85" '"14503719"
            DispLegacytoNCID = NCID_ReleaseWDutyLimitations
        
        Case "91" '"204847"
            DispLegacytoNCID = NCID_ReleasedWOLimitations
        
        Case "05"                '"540"
            DispLegacytoNCID = NCID_SickAtHome
    End Select
    
End Function

'Convert a Disposition NCID to a disposition code for the Legacy system
Private Function DispNCIDtoLegacyCode(DispNCID As String) As String

    Select Case DispNCID
        Case NCID_Admitted                  '"204849"
            DispNCIDtoLegacyCode = "95"
            
        Case NCID_Expired                   '"549"
            DispNCIDtoLegacyCode = "96"
            
        Case NCID_ImmediateReferal          '"204848"
            DispNCIDtoLegacyCode = "90"
        
        Case NCID_LeftAgainstAdvice         '"546"
            DispNCIDtoLegacyCode = "06"
        
        Case PatientLeftWithoutBeingSeen '"13"
            DispNCIDtoLegacyCode = "97"
        
        Case NCID_ReleaseWDutyLimitations   '"14503719"
            DispNCIDtoLegacyCode = "85"
        
        Case NCID_ReleasedWOLimitations     '"204847"
            DispNCIDtoLegacyCode = "91"
        
        Case NCID_SickAtHome                '"540"
            DispNCIDtoLegacyCode = "05"
    End Select
    
End Function


Public Function SaveToEncounter() As Boolean
    Dim objEnc          As CHCSIIEncounterOps.Encounter
    Dim objPersist      As IPersist
    Dim bRet            As Boolean
    Dim objAdminCode    As AdminCode
    Dim szAdminCode     As String
    Dim objAppts        As CHCSII_Appointments.Appointments
    Dim objAppt         As CHCSII_AppointmentClient.Appointment
    Dim sOldMetaData    As String
            
    
    On Error Resume Next
        
    Set objEnc = gobjEncParent.CurrentEncounter

    'Do this for all modes
    With objEnc
        
        gobjEncParent.SetStatus Me.CodedValues.DispositionStatus
        .DispositionNCID = Me.CodedValues.DispositionNCID
        .EandMNCID = Me.CodedValues.EMCodeNCID
        .PatientStatus = Me.CodedValues.AppointmentClass
        
        'SCR 65426 Don't update the appointment in inpatient encounters.
        If gDisplayMode = Default Then
            .WorkLoad = Me.CodedValues.WORK_LOAD_INDICATOR
        End If
        
        .EandMCalc = Me.CodedValues.EandMCalc
        .FollowupComments = Me.CodedValues.FollowupComments
        .DiscussedComments = Me.CodedValues.DiscussedComments
        .DBNICategory = Me.CodedValues.DBNICategory
        .DBNICause = Me.CodedValues.DBNICause
        .DispositionText = Me.CodedValues.DispositionText
        
                                                         'SCR 65426
        If gobjShared.CmdLineSwitches("GEMS_ARMY") <> "" And gDisplayMode <> InpatientDischarge Then
            'update Patient Tracking (Appointments) display
            If mAppointmentOps Is Nothing Then
                'Dim objAppts As CHCSII_Appointments.Appointments
                Set objAppts = gobjComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
                Set mAppointmentOps = objAppts.AppointmentOps
            End If
            If IsNumeric(.DBNICategory) And IsNumeric(.DBNICause) Then
                mAppointmentOps.UpdateAppointmentInjury_Illness .AppointmentID, _
                                            Me.CodedValues.DBNICategoryDesc, Me.CodedValues.DBNICauseDesc
            Else
                mAppointmentOps.UpdateAppointmentInjury_Illness .AppointmentID, "", ""
            End If
            
            'Refresh Appt Data SCR#49319
            gobjComm.Message cwiREFRESH_DATA, "APPTID|" & .AppointmentID, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_DISPOSITION_APPLICATION

        End If

        'Public Const NCID_CONSULTATION_REQUESTED    As String = "14511053"
        'Public Const NCID_CONSULTATION_DESELECTED   As String = "14511058"
        'Public Const NCID_REFERRED_TO_ANOTHER_PROVIDER As String = "14511054"
        'Public Const NCID_CONVALESCENT_LEAVE        As String = "14511055"
        'Public Const NCID_MEDICAL_BOARD             As String = "14511056"
        'Public Const NCID_MEDICAL_HOLD              As String = "14511057"
        szAdminCode = vbNullString
        For Each objAdminCode In Me.MetaData.AdminCodes
            Select Case objAdminCode.Code
            Case 1
                szAdminCode = szAdminCode & NCID_CONSULTATION_REQUESTED & "/"
            Case 2
                szAdminCode = szAdminCode & NCID_REFERRED_TO_ANOTHER_PROVIDER & "/"
            Case 3
                szAdminCode = szAdminCode & NCID_CONVALESCENT_LEAVE & "/"
            Case 4
                szAdminCode = szAdminCode & NCID_MEDICAL_BOARD & "/"
            Case 5
                szAdminCode = szAdminCode & NCID_MEDICAL_HOLD & "/"
            End Select
        Next
        
        If Len(szAdminCode) Then
            szAdminCode = Left(szAdminCode, Len(szAdminCode) - 1)
        End If
        .DispAdminOption = szAdminCode
        
        'All The Other Bits and peices are Persisted As XML
        Set objPersist = Me.MetaData
        
        sOldMetaData = .DispMetadata
        
        .DispMetadata = objPersist.Persist
        
        If StrComp(sOldMetaData, .DispMetadata, vbTextCompare) <> 0 Then
            'User changes Disposition data SCR#39592
            mbUpdated = True
        Else
            mbUpdated = False
        End If
        
        Debug.Print .DispMetadata
        
'Appt Classification:
        If objEnc.EncounterType = TeleConsult Then
            'SCR 65426
        ElseIf gDisplayMode <> InpatientDischarge Then
            Select Case Me.MetaData.AppointmentClass
            Case "Inpatient"
                objEnc.EncounterType = Inpatient
            Case "Outpatient APV"
                objEnc.EncounterType = Ambulatory
            Case "Outpatient", "Outpatient Observation", "Outpatient Home"
                objEnc.EncounterType = Outpatient
            Case Else
                objEnc.EncounterType = Outpatient
            End Select
            
        End If
    
    End With
End Function
'*******************************************************************************************
' This takes a small XML fragment and a tag name and returns the value of that tag,
' if any.
'
'*******************************************************************************************
Private Function UnLongoIzer(XML As String, TagName As String) As String
    Dim pos         As Long
    Dim temp        As String
    Dim quote1      As Long
    Dim quote2      As Long
    
    pos = InStr(1, XML, TagName, vbTextCompare)
    
    If pos = 0 Then Exit Function
    temp = Mid(XML, pos)
    
    pos = InStr(1, temp, "=")
    
    If pos Then
                                'ASCII 34 = double quote character
        quote1 = InStr(pos, temp, Chr(34)) + 1
        quote2 = InStr(quote1, temp, Chr(34))
        UnLongoIzer = Mid(temp, quote1, quote2 - (quote1))
    End If

End Function
Public Function LoadFromEncounter()

    Dim objEnc      As CHCSIIEncounterOps.Encounter
    Dim objPersist  As IPersist
    Dim bRet        As Boolean
    Dim objAppt As CHCSII_AppointmentClient.Appointment
    
    If gobjEncParent Is Nothing Then Exit Function
    
    Set objEnc = gobjEncParent.CurrentEncounter
    '<SCR 36876
    Select Case True
        
    Case gobjShared.IsAppMode(modeCDR), gobjShared.IsAppMode(modeITT), gobjShared.IsAppMode(modeTheater)
        With objEnc
        
            Me.CodedValues.DispositionStatus = .DispStatus
            Me.CodedValues.DispositionNCID = .DispositionNCID
            Me.CodedValues.EMCodeNCID = .EandMNCID
            'PERF non needed
            'Set objAppt = GetAppointment()
            Me.CodedValues.AppointmentClass = .PatientStatus
            Me.CodedValues.WORK_LOAD_INDICATOR = .WorkLoad

            Me.CodedValues.DBNICategory = .DBNICategory
            Me.CodedValues.DBNICause = .DBNICause
            
           'All The Other Bits and peices are Persisted As XML
           If Len(.DispMetadata) Then
                Set objPersist = Me.MetaData
                bRet = objPersist.Restore(.DispMetadata)
           End If
            
            'Me.MetaData.AdminCodes = .DispAdminOption   ' Delimited by "/"; Actual HDD NCID's are stored
            
        End With

    End Select
    '>SCR 36876
End Function
Public Property Get RTFs() As RTFs
    Set RTFs = mvarRTFs
End Property


Public Property Set RTFs(vData As RTFs)
    Set mvarRTFs = vData
End Property



Public Property Get MetaData() As MetaData
    Set MetaData = mvarMetaData
End Property


Public Property Set MetaData(vData As MetaData)
    Set mvarMetaData = vData
End Property


Private Sub Class_Initialize()
    'create the mCodedValues object when the DataRepository class is created
    Set mvarCodedValues = New CodedValues
  'create the mMetaData object when the DataRepository class is created
  Set mvarMetaData = New MetaData
  'create the mRTFs object when the DataRepository class is created
  Set mvarRTFs = New RTFs
End Sub


Public Property Get CodedValues() As CodedValues
    Set CodedValues = mvarCodedValues
End Property


Public Property Set CodedValues(vData As CodedValues)
    Set mvarCodedValues = vData
End Property
Private Sub Class_Terminate()
    Set mvarRTFs = Nothing
    Set mvarMetaData = Nothing
    Set mvarCodedValues = Nothing
End Sub

Public Function DispositionStatus() As EncStatus


    Dim sEMCode             As String
    Dim lEMCode             As Long
    Dim lStatus             As EncStatus
    Dim objEnc              As CHCSIIEncounterOps.Encounter
    
    Set objEnc = gobjEncParent.CurrentEncounter
    
    lStatus = EncStatus.complete  'Set default status:

    'SF 9-11-06 105696
    If InpatientDischarge Then
        GoTo XIT
    End If
    
    With Me

        'Check E&M Code
        If Len(.CodedValues.EMCode) = 0 Then
            'Set to default:
            lStatus = EncStatus.Incomplete
            GoTo XIT
        End If
        
        If Not IsNumeric(.CodedValues.EMCode) Then
            lStatus = EncStatus.Incomplete
            GoTo XIT
        End If

        lEMCode = CLng(.CodedValues.EMCode)
        If lEMCode <= 0 Then
            'Set default:
            lStatus = EncStatus.Incomplete
        End If

        Select Case gobjEncParent.CurrentEncounter.EncounterType
        
        Case Outpatient, Ambulatory
            Select Case CLng(.CodedValues.EMCode)
                Case 99201 To 99499
                Case Else
                    lStatus = EncStatus.Incomplete
            End Select
            
        Case Inpatient, TeleConsult
            Select Case CLng(.CodedValues.EMCode)
                Case 99201 To 99499
                Case Else
                    lStatus = EncStatus.Incomplete
                    
           End Select
           
        End Select
        
        'Check if DNBI data exists
        If (Len(.CodedValues.DBNICategory) = 0 Or Len(.CodedValues.DBNICause) = 0) And Len(cwobjCWComm.CmdLineSwitch("NO_DNBI")) = 0 Then
            lStatus = EncStatus.Incomplete
        End If
    
    End With 'Me
    


XIT:
    DispositionStatus = lStatus

End Function 'DispositionStatus
'New routines
Public Sub SetClinicCapabilities()
    Dim objClinic       As CHCSII_ClinicClient.Clinic
    
    If gDisplayMode = InpatientDischarge Then Exit Sub
    
    Set objClinic = modControls.GetClinic(gobjEncParent)
    If objClinic Is Nothing Then
        Debug.Print "Clinic not created for " & EncDoc().ClinicNCID
        Exit Sub
    End If
    
    If InStr(1, objClinic.LocationType, "AMBULATORY", vbTextCompare) > 0 Then
       mbAmbulatoryAuthorization = True
    Else
       mbAmbulatoryAuthorization = False

       msMEPRSCode = objClinic.MEPRSCode
    End If

    msDMISIdCode = objClinic.DmisIdCode

End Sub
'convenience function
Friend Function EncDoc() As CHCSIIEncounterOps.Encounter

    If gobjEncParent Is Nothing Then Exit Function
    
    Set EncDoc = gobjEncParent.CurrentEncounter

End Function
Public Property Get AmbulatoryAuthorization() As Boolean

    AmbulatoryAuthorization = mbAmbulatoryAuthorization

End Property
Public Property Get MEPRSCode() As String

    MEPRSCode = msMEPRSCode

End Property
Public Property Get DmisIdCode() As String

    DmisIdCode = msDMISIdCode

End Property

Public Property Get Updated() As Boolean
    Updated = mbUpdated
End Property


