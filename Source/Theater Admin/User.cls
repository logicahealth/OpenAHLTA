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
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mNCID         As String
Public ClinicNCID   As String
Public HomePhone    As String
Public WorkPhone    As String
Public EMail        As String
Public Domain       As String
Public Timer        As Long
Public ProviderName As String
Public PrimaryRole  As String
Public LOGON_NAME    As String
Public LOGON_PASSWD  As String
Public USER_INITIALS As String
Public ACTIVE_FLAG   As String
Public SPECIALTY_NCID As String

Public Signature_Class As Integer
Public SpecialyCode As Long
Public SSN          As String
Public Facilities   As Collection  ' of user_fac
Public Unit_Number  As String
Public User_Number  As String
Public PCM_Code     As String
Public IEN          As String
Public Resources    As Collection  ' of user_res
Public Roles        As Collection  ' of strings
Public OtherPrivs   As Collection  ' of privs
Public NewUser      As Boolean

Public Property Get NCID() As String
    
    NCID = mNCID
    
End Property

Public Property Let NCID(vData As String)
    
    mNCID = vData
    
End Property

Private Sub Class_Initialize()
    Set Facilities = New Collection
    Set Resources = New Collection
    Set Roles = New Collection
    Set OtherPrivs = New Collection
End Sub
