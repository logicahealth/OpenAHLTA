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
Attribute VB_Name = "DentalChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:48:17

'-------------------------------
'This is the individual representation of a child ole module.
'It contains information needed to identify and control those childres.
'The child module itself is the mobjModule object
'The rest of this object mimics the core info array.
'-------------------------------
Private mvarProgID            As Long      '- program id from core
Private mvarServerIdx         As Long      '-the index of the server/form/rtb
Private mobjModule            As Object    '- the running module as an object (from core)
Private mvarServerName        As String    ' from the index and global name
Private mvarAppNCID           As String    ' the server's ncid
Private mvarChildHeight       As Long
Private mvarChildWidth        As Long
Private mobjComm              As Object
Private mobjLogonEnvironment  As Object
Private mobjConfig            As Object
Private mobjPatient           As Object
Private mobjUser              As Object
Private mbRegularViewRunning  As Boolean   '- is this the one with the reg view running?
Private mbModuleLoaded        As Boolean   '- did the module load?
Private mbModuleOpenView      As Boolean   '- was the modules' view opened?
Private miView                As Integer
Private mlHandle              As Long
Private mcolMenuItems         As Collection

#If DEBUGON Then
Private Const dbModule = "DentalChild.cls"
#End If
Public Property Get childMenuItems() As Collection

   If mcolMenuItems Is Nothing Then
      Set mcolMenuItems = New Collection
   End If
   
   
   Set childMenuItems = mcolMenuItems

End Property

Public Property Let ModuleLoaded(RHV As Boolean)

   mbModuleLoaded = RHV

End Property

Public Property Get ModuleLoaded() As Boolean

   ModuleLoaded = mbModuleLoaded

End Property

Public Property Let parentHandle(RHV As Long)

   mlHandle = RHV

End Property

Public Property Get parentHandle() As Long

   parentHandle = mlHandle

End Property

Public Property Let childHeight(Value As Long)

   mvarChildHeight = Value

End Property

Public Property Get childHeight() As Long

   childHeight = mvarChildHeight

End Property

Public Property Let childWidth(Value As Long)

   mvarChildWidth = Value

End Property

Public Property Get childWidth() As Long

   childWidth = mvarChildWidth

End Property

Public Property Let ModuleOpenView(RHV As Boolean)

   mbModuleOpenView = RHV

End Property

Public Property Get ModuleOpenView() As Boolean

   ModuleOpenView = mbModuleOpenView

End Property

Public Property Let ViewType(RHV As Integer)

   miView = RHV

End Property

Public Property Get ViewType() As Integer

   ViewType = miView

End Property

Public Property Let RegularViewRunning(RHV As Boolean)

   mbRegularViewRunning = RHV

End Property

Public Property Get RegularViewRunning() As Boolean

   RegularViewRunning = mbRegularViewRunning

End Property

Public Property Let User(ByVal RHV As Object)

   Set mobjUser = RHV

End Property

Public Property Get User() As Object

   Set User = mobjUser

End Property

Public Property Let Patient(ByVal RHV As Object)

   Set mobjPatient = RHV

End Property

Public Property Get Patient() As Object

   Set Patient = mobjPatient

End Property

Public Property Let Config(ByVal RHV As Object)

   Set mobjConfig = RHV

End Property

Public Property Get Config() As Object

   Set Config = mobjConfig

End Property

Public Property Let LogonEnvironment(ByVal RHV As Object)

   Set mobjLogonEnvironment = RHV

End Property

Public Property Get LogonEnvironment() As Object

   Set LogonEnvironment = mobjLogonEnvironment

End Property

Public Property Let Comm(ByVal RHV As Object)

   Set mobjComm = RHV

End Property

Public Property Get Comm() As Object

   Set Comm = mobjComm

End Property

Public Property Let ServerName(ByVal RHV As String)

   mvarServerName = RHV

End Property

Public Property Get ServerName() As String

   ServerName = mvarServerName

End Property

Public Property Let AppNCID(RHV As String)

   mvarAppNCID = RHV

End Property

Public Property Get AppNCID() As String

   AppNCID = mvarAppNCID

End Property

Public Property Get ServerIndex() As Long

   ServerIndex = mvarServerIdx

End Property

Public Property Let ServerIndex(RHV As Long)

   mvarServerIdx = RHV
   mvarServerName = Sect(RHV)

End Property

Public Property Get Module() As Object

   Set Module = mobjModule

End Property

Public Property Let Module(RHV As Object)

   Set mobjModule = RHV

End Property

Public Property Let progID(ByVal RHV As Long)

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProgID = 5
   mvarProgID = RHV

End Property

Public Property Get progID() As Long

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProgID
   progID = mvarProgID

End Property

Public Sub CloseChild()

   On Error GoTo ErrHandler

#If DEBUGON Then
   Dim DR  As DebugRec

   If Not DL Is Nothing Then DL.DebugLog Exe, dbModule, "CloseChild", "Begin", DR, mvarServerName
#End If
   
'-- get rid of menu items
   
'-- tell child to close it's form
   If mobjModule Is Nothing Then _
      Exit Sub
   
   mobjModule.CloseView miView

   DoEvents
   mbModuleLoaded = False
   Set mobjModule = Nothing

   DoEvents
#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, dbModule, "CloseChild", "End", DR, mvarServerName
#End If
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.CloseView: " & ServerName & Chr(10), "Enc", vbCritical)
   Set mobjModule = Nothing
   
End Sub

Friend Function StartModule(robjDental As Dental, idx As Integer, robjDentEnc As IntrfcCHCSII_Dental.IDental_Enc) As Boolean

'-- need everything in initialize
'-- establish the server name
   Dim sProgID      As String
   Dim rhMainForm   As Long
   Dim ret          As Boolean
   Dim bOleRunning  As Boolean
   Dim oEnctr       As Object

#If DEBUGON Then
   Dim DR  As DebugRec

   If Not DL Is Nothing Then DL.DebugLog Exe, dbModule, "StartModule", "Begin", DR, "ServerName", glServerInfo(idx)
#End If
   mvarAppNCID = glServerInfo(idx)
   
   '-- If Not DebugMode Then On Error GoTo ErrHandler
   Set mobjModule = cwobjCWComm.InitializeOLEServer(glServerInfo(idx))
   sProgID = ""

   If Not mobjModule Is Nothing Then
      
'-- get the progid for this app
      Call cwobjCWComm.Message(cwiGET_APPLICATION_PROGID, sProgID, NCID_CW_APPLICATION, glServerInfo(idx))
      
      If sProgID <> "" Then
         mvarProgID = Val(sProgID)
      End If

      If TypeOf mobjModule Is ICHCSII_Dental Then
         Dim oTempDental  As ICHCSII_Dental

         Set oTempDental = mobjModule
         Call oTempDental.SetDentalParent(robjDental)
         Call oTempDental.SetEncounterOps(gobjEnc)
      End If
      
      If TypeOf mobjModule Is ICHCSII_Dental2 Then
        Dim objTempDental2 As ICHCSII_Dental2
        
        Set objTempDental2 = mobjModule
        With objTempDental2
            Set .DentalEncounter = robjDentEnc
        End With
      End If

      StartModule = True
   Else   '<bjr100501 - SCR17713
      StartModule = False
   End If                              '-- Not mobjModule Is Nothing Then

   Set oTempDental = Nothing
   Set robjDental = Nothing
   Set objTempDental2 = Nothing
   
#If DEBUGON Then
   If Not DL Is Nothing Then DL.DebugLog Exe, dbModule, "StartModule", "End", DR, "ServerName", ServerName
#End If
   
   Exit Function
   
   
ErrHandler:
   '-- -show the VB error---
   '-- Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental.StartModule: " & ServerName, vbCritical)
   
   '-- -set return value---
   StartModule = False
   Set mobjModule = Nothing
   
   Exit Function

End Function

