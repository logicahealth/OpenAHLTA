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
Attribute VB_Name = "CDMEvntSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sModuleName As String
Private lAppMode As Long

'SL- 09/21/04 - reference counting (SCR #51729)
Public ReferenceCount As Long

'Private Const gNCID_DATA_MANAGER_BACKEND As String = "1999983"
'Private Const gNCID_DATA_MANAGER_EVNTPROC_APPLICATION As String = "1999982"

Public Enum EvntBarCodeFocus
   BarCode_None = 0
   BarCode_DMFE = 1000
   BarCode_IMM = 2000
End Enum

Private BarCodeDestination As EvntBarCodeFocus

Public Event SendEvent(Status As Long, DataType As String, cData As Collection, xmlMsg As String, AppNCID As String)
Public Event BarCode(xmlMsg As String)
Public Event A28Event()

'<< Begin:SCR #33453;   Developer: Brian Mowbray 03/07/2003 02:06 PM
Public Event CAC(XML As String)
'>> End: SCR #33453;

'<< SCR #33453;   Developer: Brian Mowbray 03/07/2003 02:07 PM
Public Sub SendCACMessage(XML As String)
    RaiseEvent CAC(XML)
End Sub

'<< SCR #26534;   Developer: Brian Mowbray 08/12/2002 02:50 PM
Public Sub TerminateEvntSvr(Optional Immediately As Boolean = False)

On Error Resume Next
  
    If Immediately Then
        ReferenceCount = 1
    End If

    'SCR #51729 - close if no more references
    ReferenceCount = ReferenceCount - 1
    
    'SF SCR 71407 6/6/05 'Close this thing down
    'mcctConnector.CloseIfUnused
    If ReferenceCount = 0 Then
        ShutdownServer
    End If
    'SF SCR 71407 6/6/05
    
End Sub

'<< SCR #28900;   Developer: Brian Mowbray 12/18/2002 05:20 PM
Public Sub CloseCommPort()
  On Error Resume Next
  gbCommPortOn = False
  SetCommPort
End Sub
'<< SCR #28900;   Developer: Brian Mowbray 12/18/2002 05:20 PM
Public Sub OpenCommPort()
  On Error Resume Next
  GetCommPortRegValue enCommPortAll
  ' Begin:SCR #39295;   Developer: MAG 07/29/2003 03:08 PM
  SetCommPort
  ' End: SCR #39295;
End Sub

'<< SCR #25934;   Developer: Brian Mowbray 08/01/2002 05:14 PM
Public Property Get AppMode() As Long
  AppMode = lAppMode
End Property



'<< SCR #25934;   Developer: Brian Mowbray 08/01/2002 05:14 PM
Public Property Let AppMode(lModeIn As Long)
  lAppMode = lModeIn
End Property
Private Sub Class_Initialize()
  
  On Error Resume Next
  
  
    WriteDebugLog "Class CMDEvntSvr.Initialized invoked"
    ReferenceCount = 0
   
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   WriteDebugLog "CDMEvntSvr.Terminate invoked"
End Sub


Public Sub SetBarcodeFocus(BCDest As EvntBarCodeFocus)
   BarCodeDestination = BCDest
End Sub

Public Sub BarCodeHandler(DataType As String, BCMsg As String)
   Dim ErrorMsg As String
   RaiseEvent BarCode(BCMsg)
   Select Case BarCodeDestination
      Case BarCode_DMFE
        RaiseEvent SendEvent(0, DataType, DMFECollection, BCMsg, NCID_DATA_MANAGER)
      Case BarCode_IMM
        RaiseEvent SendEvent(0, DataType, IMMCollection, BCMsg, NCID_IMMUNIZATION_SYNCH_APPLICATION)
   End Select
End Sub


Private Function ReadDataFile(DataPath As String) As String
   Dim fso As FileSystemObject
   Dim Tobj As TextStream
   
   Set fso = New FileSystemObject
   Set Tobj = fso.OpenTextFile(DataPath, ForReading, False, TristateFalse)
   ReadDataFile = Tobj.ReadAll
   Set fso = Nothing
   Set Tobj = Nothing
End Function

Public Sub RaiseA28Event()
  RaiseEvent A28Event
End Sub

