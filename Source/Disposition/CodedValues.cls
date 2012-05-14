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
Attribute VB_Name = "CodedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public DispositionStatus As String
Public DispositionNCID As String
Public CheifComplaintCode As String
Public OnsetDate As Date
Public WorkRelated As Boolean
Public WORK_LOAD_INDICATOR As String
Public AppointmentClass As String
Public DispCode As String 'IBWA
Public EandMCalc As String
Public FollowupTimeFrame As String
Public FollowupComments As String
Public DiscussedComments As String
Public ItemsDiscussed As String
Public DispositionText As String
Public DBNICategory As String
Public DBNICause As String
Public DBNICategoryDesc As String
Public DBNICauseDesc As String

Dim m_szEMCode As String
Dim m_EMCodeNCID As String

Public Property Get EMCode() As String
    EMCode = m_szEMCode
End Property
Public Property Let EMCode(ByVal NewValue As String)
    
    Dim rs As ADODB.Recordset
    
    m_szEMCode = NewValue
    
    'Also fixes SCR-32005, 31216
    If m_szEMCode = vbNullString Then
        m_EMCodeNCID = vbNullString 'SCR-34812
        Exit Property
    End If
    
    Set rs = GetEMTable
    rs.Find "EM_CODE='" & NewValue & "'", , adSearchForward, adBookmarkFirst
    
    If rs.EOF Then
        m_EMCodeNCID = vbNullString
    Else
        m_EMCodeNCID = rs.Fields("NCID").Value
    End If
    
End Property

Public Property Get EMCodeNCID() As String
    EMCodeNCID = m_EMCodeNCID
End Property
Public Property Let EMCodeNCID(ByVal NewValue As String)

    Dim rs As ADODB.Recordset
    
    m_EMCodeNCID = NewValue
    
    If Len(m_EMCodeNCID) Then
        Set rs = GetEMTable
    'R. Foster SCR #48177  Developer: Robert Foster 01/20/2004 06:14 PM
    'R. Foster Note: Function Changed: CodedValues.cls - EMCodeNCID (Let); Profile.cls - Save;  modControls.bas - GetEMTable;      Impact of Change:  Modifiy code to account for ITT mode
    'ITT CONVERGENCE MODIFICATION
    If gobjShared.IsAppMode(modeITT) Then
        rs.Find "NCID=" & val(NewValue) & "", , adSearchForward, adBookmarkFirst
    Else
        rs.Find "NCID='" & NewValue & "'", , adSearchForward, adBookmarkFirst
    End If

        
        If rs.EOF Then
            m_szEMCode = vbNullString
        Else
            m_szEMCode = rs.Fields("EM_CODE").Value
        End If
    Else
        m_szEMCode = vbNullString
    End If

End Property

