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
Attribute VB_Name = "AnPOLEServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class represents a OLE Server that is displayed by AnP.

Option Explicit

Private mModuleID As enumAnPModules
Private mNCID As String
Private mServerName As String
Private mViewType As Integer
Private mPriv As String

Private mHeight As Long         'SCR 35637  5/13/03 Phil Crowder
Private mTop As Long            'SCR 35637  5/13/03 Phil Crowder
Private mWidth As Long          'SCR 35637  5/13/03 Phil Crowder

Private mServer As Object

Public Property Get Height() As Long
  Height = mHeight
End Property

Public Property Let Height(ByVal NewValue As Long)
  mHeight = NewValue
End Property

Public Property Get ModuleID() As enumAnPModules
  ModuleID = mModuleID
End Property

Public Property Let ModuleID(ByVal NewValue As enumAnPModules)
  mModuleID = NewValue
End Property

Public Property Get NCID() As String
  NCID = mNCID
End Property

Public Property Let NCID(ByVal NewValue As String)
  mNCID = NewValue
End Property

Public Property Get Privilege() As String
  Privilege = mPriv
End Property

Public Property Let Privilege(ByVal NewValue As String)
  mPriv = NewValue
End Property

Public Property Get Server() As Object
  Set Server = mServer
End Property

Public Property Get ServerName() As String
  ServerName = mServerName
End Property

Public Property Let ServerName(ByVal NewValue As String)
  mServerName = NewValue
End Property

Public Property Get Top() As Long
  Top = mTop
End Property

Public Property Let Top(ByVal NewValue As Long)
  mTop = NewValue
End Property

Public Property Get ViewType() As Integer
  ViewType = mViewType
End Property

Public Property Let ViewType(ByVal NewValue As Integer)
  mViewType = NewValue
End Property

Public Property Get Width() As Long
  Width = mWidth
End Property

Public Property Let Width(ByVal NewValue As Long)
  mWidth = NewValue
End Property

Public Sub CloseView()
  On Error Resume Next

  If Not (mServer Is Nothing) Then
    mServer.CloseView mViewType
  End If
  
End Sub

Public Function InitView() As Boolean
Dim AnpInterface As ICHCSII_AnP

  On Error GoTo ErrHandler
         
  If Not (mServer Is Nothing) Then
    InitView = True
    Exit Function
  End If
  
  Set mServer = gBB.ReadModule(mNCID)
    
  If Not (mServer Is Nothing) Then
    If TypeOf mServer Is ICHCSII_AnP Then
      Set AnpInterface = mServer
      AnpInterface.SetAnPParent gBB.AnPChildInterface
    End If
  End If
    
  InitView = Not (mServer Is Nothing)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServer.InitView", "A/P", vbExclamation

End Function

Public Function OpenView(ByVal ContainerHwnd As Long) As Boolean
Dim WndMainForm As Long

  On Error GoTo ErrHandle
  
  If Not (Server Is Nothing) Then
    Server.OpenView mViewType, ContainerHwnd, WndMainForm
    OpenView = True
  Else
    OpenView = False
  End If
  
  Exit Function
  
ErrHandle:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServer.OpenView", "A/P", vbExclamation
  
End Function

Public Function Message(ByVal MessageType As Integer, ByRef MessageText As String, _
                   ByVal SourceNCID As String) As Boolean
                   
  On Error GoTo ErrHandle
  
  If Not (mServer Is Nothing) Then
    mServer.Message MessageType, MessageText, SourceNCID
  End If
                   
  Exit Function
  
ErrHandle:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServer.Message", "A/P", vbExclamation
                   
End Function

Public Function ResizeView() As Boolean

  On Error GoTo ErrHandle
  
  If Not (Server Is Nothing) Then
    Server.ResizeView mViewType, 0, mTop, mWidth, mHeight
    ResizeView = True
  Else
    ResizeView = False
  End If
  
  Exit Function
  
ErrHandle:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServer.ResizeView", "A/P", vbExclamation
  
End Function

Public Function UserHasPriv(ByVal PrivType As CHCSII_PrivilegeType) As Boolean
Dim User As Object
  
  Set User = gBB.User
  UserHasPriv = User.HasPrivilegeEx(mPriv, PrivType)
  
End Function
