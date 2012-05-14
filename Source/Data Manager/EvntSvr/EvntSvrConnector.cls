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
Attribute VB_Name = "EvntSvrConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get EvntSvr() As CDMEvntSvr
    ' If the shared EvntSvr object hasn't been
    '   created, create it and store a global reference
    '   to it.
    
    If gEvntSvr Is Nothing Then
        WriteDebugLog "Creating instance of CDMEvntSvr Class"
        Set gEvntSvr = New CDMEvntSvr
    End If
    
    'core is keeping the reference once this is called, and without
    '   changes to core, there's no other good way to count references.
    'each call to this method must be matched by a call to TerminateServer,
    '   or EventServer will not know when to close
    gEvntSvr.ReferenceCount = gEvntSvr.ReferenceCount + 1
    
    
    Set EvntSvr = gEvntSvr
End Property

Private Sub Class_Initialize()
'   If gEvntSvr Is Nothing Then
'      Set gEvntSvr = New CDMEvntSvr
'   End If
    WriteDebugLog "Class EvntSvrConnector.Initialized invoked"
End Sub

Private Sub Class_Terminate()
  WriteDebugLog "Class EvntSvrConnector.Terminate invoked"
End Sub

Public Sub CloseIfUnused()
    
    If gEvntSvr.ReferenceCount < 1 Then
    
        On Error Resume Next
        
        ShutdownServer
    End If
End Sub

