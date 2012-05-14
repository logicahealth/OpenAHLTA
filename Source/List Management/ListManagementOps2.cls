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
Attribute VB_Name = "ListManagementOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"OtherClinics"
Option Explicit

Private mvarUser As ListOwner
Private mvarFacility As ListOwner
Private mvarClinic As ListOwner
Private mvarGUIListOwners As ListOwners
Private mvarListOwners As ListOwners
Private mbNewAppt As Boolean 'Luis Ibarra
Private mvarProviderNCID As String 'SCR 130196


Public Property Get ProviderNCID() As String
    
    ProviderNCID = mvarProviderNCID
    
End Property

Public Property Let ProviderNCID(vdata As String)
    'SCR 130196
    mvarProviderNCID = vdata
    
End Property

Public Function SetNewClinicFavorites(ClinicNCID As String) As Boolean
    
    Dim APValue As String
    Dim oGemsDas As ICHCSII_DataAccess.ICHCSII_DAS
    Dim oConn As CHCSII_CONN.Conn                   'hwcc, convert to use CHCSII_CONN
    Dim rs As ADODB.Recordset
    Dim SQLText As String

    Const FacilityListKey = "ClinicDefault"
    
    On Error GoTo errorHandler
    
    If gobjShared Is Nothing Then Set mobjshared = New CWShared
    If gobjShared.CmdLineSwitches.Item("GEMS_ARMY") = "" Then Exit Function 'if this is not Triage then get out
    
    Set oConn = New CHCSII_CONN.Conn
    Set oGemsDas = oConn.CHCSII_DAS
    
    SQLText = "Select apValue from OSH_INI_Settings" _
        & " where Application_NCID = " & 186412 _
        & " AND apSection = '" & COMPLAINT_LIST_NAME _
        & "' AND apKey = '" & "ClinicDefault" & "'" _
        & " AND apOwnerNCID = " & 0
            
    Set rs = oGemsDas.OpenRecordset(SQLText)

    cwobjConfig.SetValue NCID_LIST_MGMT_APPLICATION, _
                            cwobjLogonEnviron.FacilityNCID, _
                            Trim(ClinicNCID) & " " & COMPLAINT_LIST_NAME, _
                            FacilityListKey, _
                            rs.Fields("apvalue").Value

    cwobjConfig.SaveChanges NCID_LIST_MGMT_APPLICATION, cwobjLogonEnviron.FacilityNCID
    SetNewClinicFavorites = True
    GoTo CleanUp
errorHandler:
    Call mobjshared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListManagementOps.SetNewClinicFavorites", "ListManagement", vbExclamation)
    'Resume Next
CleanUp:
    Set oGemsDas = Nothing
    Set oConn = Nothing
    
End Function
Public Property Get GUIListOwners() As ListOwners
    
    'SCR 101234 - LCI - 09/19/2006
    If NewAppt Then
     Set mvarGUIListOwners = Nothing
    End If
    
    If mvarGUIListOwners Is Nothing Then
    
        Set mvarGUIListOwners = New ListOwners
        'SCR 130196 Add the passed in provider ncid
        mvarGUIListOwners.Init True, True, True, mvarProviderNCID
        If Not Me.Clinic Is Nothing Then mvarGUIListOwners.Add Me.Clinic, Me.Clinic.OwnerNCID
        If Not Me.User Is Nothing Then mvarGUIListOwners.Add Me.User, Me.User.OwnerNCID
        
    End If

    Set GUIListOwners = mvarGUIListOwners
    
End Property

Public Property Get Clinic() As ListOwner
        
    If mvarClinic Is Nothing Then
        If Not gobjLogon.NursingDivisionNCID = "" Then
            Set mvarClinic = New ListOwner
            mvarClinic.Init otClinic, gobjLogon.NursingDivisionNCID, gobjLogon.FacilityNCID
        End If
    ElseIf Not CStr(mvarClinic.OwnerNCID) = gobjLogon.NursingDivisionNCID And Not gobjLogon.NursingDivisionNCID = "" Then
        Set mvarClinic = New ListOwner
        mvarClinic.Init otClinic, gobjLogon.NursingDivisionNCID, gobjLogon.FacilityNCID
    End If

    Set Clinic = mvarClinic
    
    'SCR 101234 - LCI - 09/19/2006
    If NewAppt Then
        Set mvarClinic.mobjComplaints = Nothing
    End If
    
End Property

Public Property Get Facility() As ListOwner
    If mvarFacility Is Nothing Then
        Set mvarFacility = New ListOwner
        mvarFacility.Init otFacility, gobjLogon.FacilityNCID, gobjLogon.FacilityNCID
    End If


    Set Facility = mvarFacility
End Property

Public Property Get User() As ListOwner
    If mvarUser Is Nothing Then
        Set mvarUser = New ListOwner
        mvarUser.Init otUser, gobjLogon.UserNCID, gobjLogon.UserNCID
    End If


    Set User = mvarUser
End Property

Private Sub Class_Terminate()
    
On Error Resume Next

    'Release All External references
    'Forces gobjConfig.setvalues to be called
    Call Save
    
    Set mvarGUIListOwners = Nothing
    Set mvarClinic = Nothing
    Set mvarFacility = Nothing
    Set mvarUser = Nothing

End Sub


Friend Sub Save()
    
    Dim objOwner As ListOwner
    Dim bRet As Boolean
    Dim lNCID As Long
    Dim sz As String
    Dim bSave As Boolean
    
    On Error GoTo xitErr
    
    'User/Clinic Might are part of the GUIListOwners collection, it it exists
    If Not mvarGUIListOwners Is Nothing Then
        For Each objOwner In mvarGUIListOwners
            'Debug.Assert objOwner.OwnerNCID = objOwner.SaveOwnerNCID
            If objOwner.Save Then
                bRet = gobjConfig.SaveChanges(NCID_LIST_MGMT_APPLICATION, objOwner.SaveOwnerNCID)
            End If
        Next
    Else
    'Thee GUIListOwners collection might not exist
        If Not mvarUser Is Nothing Then
            If mvarUser.Save Then
                bRet = gobjConfig.SaveChanges(NCID_LIST_MGMT_APPLICATION, mvarUser.SaveOwnerNCID)
            End If
        End If
        If Not mvarClinic Is Nothing Then
            If mvarClinic.Save Then
                bRet = gobjConfig.SaveChanges(NCID_LIST_MGMT_APPLICATION, mvarClinic.SaveOwnerNCID)
            End If
        End If
    
    End If
    
    If Not mvarFacility Is Nothing Then
        lNCID = gobjLogon.FacilityNCID
        If mvarFacility.Save Then
            bRet = gobjConfig.SaveChanges(NCID_LIST_MGMT_APPLICATION, mvarFacility.SaveOwnerNCID)
        End If
    End If
    
    Exit Sub
    
xitErr:

    sz = "List Management Config Save SaveChanges for OwnerNCID: %1 Failed."
    sz = Replace(sz, "%1", lNCID)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "ListManagementOps.Terminate", App.Title, vbExclamation
    
    Exit Sub

End Sub
Public Sub SaveUserList()
'SCR-50350 Added save user list so A/P modules make sure their data is pushed into the OSH
    
    Dim lNCID As Long
    Dim sz As String
    Dim bSave As Boolean
    
    On Error GoTo xitErr
    'The User collection might not exist
    If Not mvarUser Is Nothing Then
       If mvarUser.Save Then
          gobjConfig.SaveChanges NCID_LIST_MGMT_APPLICATION, mvarUser.SaveOwnerNCID
       End If
    End If
    Exit Sub
    
xitErr:

    sz = "List Management Config Save SaveChanges for OwnerNCID: %1 Failed."
    sz = Replace(sz, "%1", lNCID)
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION & vbCrLf & sz, "ListManagementOps.SaveUserList", App.Title, vbExclamation
    
End Sub

'SCR 101234 - LCI - 09/19/2006
Public Property Get NewAppt() As Boolean
    NewAppt = mbNewAppt
End Property

Public Property Let NewAppt(bNewAppt As Boolean)
    mbNewAppt = bNewAppt
End Property

