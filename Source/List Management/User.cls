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
Attribute VB_Name = "ListOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
'SCR 101234 - LCI - 09/19/2006
Public mobjComplaints As MMMListManagement.Entries 'local copy
Private mobjDiagnosis As MMMListManagement.Entries 'local copy
Private mobjEandMs As MMMListManagement.Entries 'local copy
Private mobjProcedures As MMMListManagement.Entries 'local copy

Public Enum EnumOwnerType
    otUser
    otClinic
    otFacility
End Enum

Public Enum EnumListType
    ltDiagnosis
    ltComplaints
    ltProcedures
    ltEandMs
End Enum

Private mvarOwnerType As EnumOwnerType
Private mlOwnerNCID As Long
Private mlSaveOwnerNCID As Long

Const UserListKey = "UserDefault"
Const FacilityListKey = "ClinicDefault"

Private mbAllowRead As Boolean
Private mbAllowWrite As Boolean
Private mbAllowDelete As Boolean
Public Property Get AllowRead() As Boolean
    AllowRead = mbAllowRead
End Property
Public Property Get AllowWrite() As Boolean
    AllowWrite = mbAllowWrite
End Property
Public Property Get AllowDelete() As Boolean
    AllowDelete = mbAllowDelete
End Property


Private Function SubKey() As String


    Select Case mvarOwnerType
        Case otUser
            SubKey = UserListKey
        Case otClinic
            SubKey = FacilityListKey
        Case otFacility
            SubKey = FacilityListKey
    End Select

End Function
Private Function Key(listType As EnumListType) As String

    Select Case listType
        Case ltDiagnosis
             Key = Trim(mlOwnerNCID) & " " & DIAGNOSIS_LIST_NAME 'UserNCID
        Case ltComplaints
             Key = Trim(mlOwnerNCID) & " " & COMPLAINT_LIST_NAME
        Case ltProcedures
             Key = Trim(mlOwnerNCID) & " " & PROCEDURE_LIST_NAME 'UserNCID
        Case ltEandMs
             Key = Trim(mlOwnerNCID) & " " & EANDM_LIST_NAME 'UserNCID
    End Select

End Function

Friend Sub Init(ByVal OwnerType As EnumOwnerType, ByVal OwnerNCID As Long, ByVal SaveOwnerNCID As Long)

    mvarOwnerType = OwnerType
    mlOwnerNCID = OwnerNCID
    mlSaveOwnerNCID = SaveOwnerNCID
    
    Call SetSecurityState
    
End Sub
Public Property Get Procedures() As MMMListManagement.Entries
    
    
    If mobjProcedures Is Nothing Then
           
        Set mobjProcedures = New MMMListManagement.Entries
        mobjProcedures.Retrieve Key(ltProcedures), SubKey(), mlOwnerNCID, _
                        mlSaveOwnerNCID, mbAllowRead, mbAllowWrite, mbAllowDelete
    End If
    
    Set Procedures = mobjProcedures
    
End Property

Public Property Get EandMs() As MMMListManagement.Entries

    If mobjEandMs Is Nothing Then
        
        Set mobjEandMs = New MMMListManagement.Entries
        mobjEandMs.Retrieve Key(ltEandMs), SubKey(), mlOwnerNCID, _
                        mlSaveOwnerNCID, mbAllowRead, mbAllowWrite, mbAllowDelete
    End If
    
    Set EandMs = mobjEandMs
    
End Property
Public Property Get SaveOwnerNCID() As Long
    SaveOwnerNCID = mlSaveOwnerNCID
End Property

Public Property Get Diagnosis() As MMMListManagement.Entries

    If mobjDiagnosis Is Nothing Then
        
        Set mobjDiagnosis = New MMMListManagement.Entries
        mobjDiagnosis.Retrieve Key(ltDiagnosis), SubKey(), mlOwnerNCID, _
                        mlSaveOwnerNCID, mbAllowRead, mbAllowWrite, mbAllowDelete
    End If
    
    Set Diagnosis = mobjDiagnosis
    
End Property


Public Property Get Complaints() As MMMListManagement.Entries

    If mobjComplaints Is Nothing Then
            
        Set mobjComplaints = New MMMListManagement.Entries
        mobjComplaints.Retrieve Key(ltComplaints), SubKey(), mlOwnerNCID, _
                        mlSaveOwnerNCID, mbAllowRead, mbAllowWrite, mbAllowDelete
        
    End If
    
    Set Complaints = mobjComplaints
    
End Property



 '{{BOOKMARK Verify Security Rules For This Module}}
Private Sub SetSecurityState()

    Dim lRet As VbMsgBoxResult
    
    mbAllowRead = False
    mbAllowWrite = False
    mbAllowDelete = False
    
    Select Case mvarOwnerType
        Case otUser
            mbAllowRead = True
            mbAllowWrite = True
            mbAllowDelete = True
        
        Case otClinic
            If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Read) Then mbAllowRead = True
            If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Write) Then mbAllowWrite = True
            If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Delete) Then mbAllowDelete = True
            
            If Not mbAllowRead Then
                If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Break_The_Glass) Then
                    lRet = MsgBox("You will be audited if you wish to proceed. Proceed?", vbYesNo + vbCritical, App.Title)
                    If lRet = vbYes Then
                        Set gobjAudit = New Audit
                        gobjAudit.TurnAuditOn 'Must turn auditing on
                        mbAllowRead = True
                        mbAllowWrite = True
                        mbAllowDelete = True
                    End If
                End If
            End If
                
        Case otFacility
            If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Read) Then mbAllowRead = True
            If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Write) Then mbAllowWrite = True
            If gobjUser.HasPrivilegeEx(Priv_Clinic_List, Priv_Delete) Then mbAllowDelete = True
    
    End Select
    


End Sub


Public Property Get OwnerNCID() As Long
    OwnerNCID = mlOwnerNCID
End Property

'Return True if any save was true (SetValue was called)
Friend Function Save() As Boolean
    
    Dim bRet As Boolean
    
    If Not mobjComplaints Is Nothing Then
        bRet = mobjComplaints.Save
    End If
    
    If Not mobjDiagnosis Is Nothing Then
        bRet = bRet Or mobjDiagnosis.Save
    End If
    
    If Not mobjEandMs Is Nothing Then
        bRet = bRet Or mobjEandMs.Save
    End If
    
    If Not mobjProcedures Is Nothing Then
        bRet = bRet Or mobjProcedures.Save
    End If

    Save = bRet
    
End Function
