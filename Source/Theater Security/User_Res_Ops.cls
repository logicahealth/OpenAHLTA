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
Attribute VB_Name = "User_Res_Ops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolRolePriv As Collection
Public colRoles As Collection
Public colPriv As Collection
Public colResources As Collection
Public colUsers As Collection


Public Function display() As Boolean
    Dim vRole As Variant
    Set frmMain.objParent = Me
    
    Dim oUser As TheaterSecurity.User
    frmMain.mbLoading = True
    For Each oUser In colUsers
        frmMain.lstUsers.AddItem oUser.ProviderName
        frmMain.lstUsers.ItemData(frmMain.lstUsers.NewIndex) = oUser.NCID
    Next oUser
    frmMain.mbLoading = False

End Function

Public Function DisplayUserAttributes(ByVal sUserNCID As String, ByRef roList As Control)
'    Dim oUser As ICHCSII_User
'    Set oUser = colUsers(sUserNCID)
'    Dim i As Integer
'    Dim sName As String
'    Dim sValue As String
'
'    For i = 1 To oUser.GetAttributeNames.count
'        sName = oUser.GetAttributeNames.Item(i)
'        sValue = oUser.GetAttribute(sName)
'        roList.AddItem sName & "=" & sValue
'    Next i
End Function
