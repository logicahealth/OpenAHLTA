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
Attribute VB_Name = "Comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'  Class Name: Comm
'
'  Description: This class acts as a "middle man" between CW and other CW OLE compatibable servers.  It is used
'               to handle the communication needs between the CW Core and other CW OLE compatiable servers.
'
'  Properties:
'    none
'
'  Methods:
'    Connect - Allows CW to connect to this OLE Server and pass in the "real" communication object used to
'              communicate with CW.
'    AddPopupMenuToCW - Allows an application to add a top-level menu to the main CW menu.
'    AddPopupMenu - Allows an application to add a pop-up menu to an existing menu.
'    AddMenuItem - Allows an application to add a menu item to an existing menu.
'    UpdateMenuItem - Allows an application to "update" a menu item, i.e. enable/disable,
'                     check/uncheck.
'    RefreshMenu - Causes the main CW menu to be "redrawn".
'    Message - Allows for a common "message interface" between CW and CW OLE compatiable servers for
'              miscellaneous messages.
'    CloseView - Allows an OLE server to "close a view".
'    StartOLEServer - Used by a child to open a sub child
'*****************************************************************************************************************
Option Explicit


'****************************************************************************
'   V A R I A B L E S     S E C T I O N
'****************************************************************************
'---communication object used to communicate with CW---
Private mCWComm As Object

'---variables used to store object property values---
Private msReportsPath As String
Private m_sHelpFilePath As String
Private msRTDataPath As String
Private moMainWindow As Object
Private mlOEMedsState As Long     'Solves Out-Of-Body problem w/ OE Meds

'---needed for conversion from local to gmt and back
Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private tzInfo As TIME_ZONE_INFORMATION
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)


Public Function AddActionMenuItem(ByVal vsMenuText As String, _
                                  ByVal vsBitMap As String, _
                                  ByVal vlOwnerProgId As Long, _
                                  ByVal vsToolTip As String) As Long
                                  
    AddActionMenuItem = mCWComm.AddActionMenuItem(vsMenuText, _
                                                      vsBitMap, _
                                                      vlOwnerProgId, _
                                                      vsToolTip)
End Function
Public Function AddListenerForMessage(DestinationServer As Object, ByVal SourceServerNCID As String, ByVal MessageType As Integer, ByVal DestinationServerNCID As String) As Boolean
    AddListenerForMessage = mCWComm.AddListenerForMessage(DestinationServer, SourceServerNCID, MessageType, DestinationServerNCID)
End Function
Public Function AddPrintMenuItem(ByVal vsMenuText As String, _
                                 ByVal vlOwnerProgId As Long, _
                                 ByVal vsToolTip As String) As Long
                                 
    AddPrintMenuItem = mCWComm.AddPrintMenuItem(vsMenuText, _
                                                    vlOwnerProgId, _
                                                    vsToolTip)
End Function
Public Function AddPrintPreviewMenuItem(ByVal vsMenuText As String, _
                                        ByVal vlOwnerProgId As Long, _
                                        ByVal vsToolTip As String) As Long
                                 
    AddPrintPreviewMenuItem = mCWComm.AddPrintPreviewMenuItem(vsMenuText, _
                                                           vlOwnerProgId, _
                                                           vsToolTip)
End Function

Public Function AddViewMenuItem(ByVal vsMenuText As String, _
                                ByVal vlOwnerProgId As Long, _
                                ByVal vsToolTip As String) As Long
                                
    AddViewMenuItem = mCWComm.AddViewMenuItem(vsMenuText, _
                                                  vlOwnerProgId, _
                                                  vsToolTip)
End Function

Public Property Get BroadcastMessageHandler() As BroadcastMessageHandler
    Set BroadcastMessageHandler = mCWComm.GetBroadcastMessageHandler()
End Property

'************************************************************************************************************
'  Description: This subroutine calls the corresponding method in the "real" CW communcation object which
'               does the "cleanup work" required when an OLE Server closes a view from within CW.
'
'  Parameters:
'    viViewType - Contains the "type" of view that is being closed.
'    vlProgId   - Contains the "ProgId" of the OLE server that is closing a view from within CW.
'************************************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer, ByVal vlProgId As Long)
    On Error GoTo ErrHandler
    mCWComm.CloseView viViewType, vlProgId
    
    Exit Sub
    
ErrHandler:
    Call gShowVBError(Err.Number, Err.Description, "Comm.CloseView", vbExclamation)
End Sub
Public Property Get CmdLineSwitch(ByRef Key As String) As String
    CmdLineSwitch = mCWComm.CmdLineSwitch(Key)
End Property
'******************************************************************************************************
'  Description: This function is called when CW wants to "connect" to this OLE Server.
'
'  Parameters:
'    robjCWComm - Contains the communication object used to communication with CW.
'
'  Return Value:
'    Returns TRUE if the connection was successful.
'    Returns FALSE if an error occurred doing the connection.
'******************************************************************************************************
Public Function Connect(robjCWComm As Object) As Boolean
    On Error GoTo ErrHandler
    
    '---save a reference to the CW communication object---
    Set mCWComm = robjCWComm
    Connect = True
    Exit Function
    
ErrHandler:
    Call gShowVBError(Err.Number, Err.Description, "Comm.Connect", vbCritical)
    Connect = False
End Function

Public Function DeleteMenuItem(ByVal vlMenuId As Long)

    DeleteMenuItem = mCWComm.DeleteProgMenuItem(vlMenuId)

End Function

'******************************************************************************************
'  Description: This subroutine calls the corresponding method in the "real" CW communcation
'               object which is used for "extended" communication purposes between CW and CW
'               OLE compatible servers.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    rsMessage - Contains the message text to process.  Depending on the message "type", this value could be
'                modified upon return.
'    vsDestinationServerNCID - Contains the NCID of the destination OLE Server to send the message to.
'    vsSourceServerNCID - Contains the NCID of the source OLE Server that sent the message.
'    vlFlagBits - Contains flag bits set for special purposes which are "message specific".
'
'  Return Value:
'    Returns TRUE or FALSE depending upon the message "type".  The default return value is
'      TRUE.
'******************************************************************************************
Public Function MessageEx(ByVal viMessageType As Integer, _
                                rsMessage As String, _
                                ByVal vsDestinationServerNCID As String, _
                                ByVal vsSourceServerNCID As String, _
                                ByVal vlFlagBits As Long) As Boolean
    On Error GoTo ErrHandler
        
    MessageEx = mCWComm.MessageEx(viMessageType, _
                                      rsMessage, _
                                      vsDestinationServerNCID, _
                                      vsSourceServerNCID, _
                                      vlFlagBits)
    Exit Function
    
ErrHandler:
    Call gShowVBError(Err.Number, Err.Description, "Comm.MessageEx", vbExclamation)
    MessageEx = False
End Function


'*****************************************************************************************************************
'  Description: This subroutine calls the corresponding method in the "real" CW communcation object which is
'               used for miscellaneous communication purposes between CW and CW OLE compatible servers.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    rsMessage     - Contains the message text to process.  Depending on the message "type", this value could be
'                modified upon return.
'    vsDestinationServerNCID - Contains the NCID of the destination OLE Server to send the message to.
'    vsSourceServerNCID      - Contains the NCID of the source OLE Server that sent the message.
'    bStartOLEServer         - By default, CW will launch a module before handing it the message. This allows a 'quiet send'.
'
'  Return Value:
'    Returns TRUE or FALSE depending upon the message "type".  The default return value is
'      TRUE.
'*****************************************************************************************************************
Public Function Message(ByVal viMessageType As Integer, _
                              rsMessage As String, _
                        ByVal vsDestinationServerNCID As String, _
                        ByVal vsSourceServerNCID As String, _
               Optional ByVal bStartOLEServer = True) As Boolean
               
    On Error GoTo ErrHandler
'<< SCR 11207 - allow modules to send quite messages (don't start the server if it's not running)
    '---call the corresponding method in the "real" CW communication object---
    Message = mCWComm.Message(viMessageType, rsMessage, vsDestinationServerNCID, vsSourceServerNCID, bStartOLEServer)
    
    Exit Function
    
ErrHandler:
    Call gShowVBError(Err.Number, Err.Description, "Comm.Message" & vbCrLf & "MessageType=" & viMessageType & vbCrLf & "Message=" & rsMessage & vbCrLf & "Destination=" & vsDestinationServerNCID & vbCrLf & "Source=" & vsSourceServerNCID, vbExclamation)
    Message = False
End Function

'***************************************************************************************************************
'  Description: This function calls the corresponding method in the "real" CW communcation object which updates
'               the specified menu item according to the action, "vsUpdateAction".
'
'               Supported update actions include:
'                    "enable"
'                    "disable"
'                    "check"
'                    "uncheck"
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item to update.
'    vsUpdateAction - Contains the "update action" to perform on the specified menu item.
'
'  Return Value:
'    Returns TRUE if the menu item was updated successfully.
'    Returns FALSE if an error occurred in updating the menu item.
'***************************************************************************************************************
Public Function UpdateMenuItem(ByVal vlMenuId As Long, ByVal vsUpdateAction As String) As Boolean
    On Error GoTo ErrHandler
        
    '---call the corresponding method in the "real" CW communication object---
    UpdateMenuItem = mCWComm.UpdateProgMenuItem(vlMenuId, vsUpdateAction)
    Exit Function
    
ErrHandler:
    Call gShowVBError(Err.Number, Err.Description, "Comm.UpdateMenuItem", vbExclamation)
    UpdateMenuItem = False
End Function


Private Sub Class_Terminate()
    Set mCWComm = Nothing
End Sub

Public Function GetStatusText() As String:    GetStatusText = mCWComm.GetStatusText(): End Function

Public Sub SetStatusText(Text As String, Optional Seconds As Variant)
    Call mCWComm.SetStatusText(Text, Seconds)
End Sub

Public Property Get ReportsPath() As String
    ReportsPath = msReportsPath
End Property
Public Property Let ReportsPath(ByVal vNewValue As String)
    msReportsPath = vNewValue
End Property
Public Property Get HelpFilePath() As String
    HelpFilePath = m_sHelpFilePath
End Property
Public Property Let HelpFilePath(ByVal vsNewValue As String)
    m_sHelpFilePath = vsNewValue
End Property
Public Property Get RTDataPath() As String
    RTDataPath = msRTDataPath
End Property
Public Property Let RTDataPath(ByVal vsNewValue As String)
    msRTDataPath = vsNewValue
End Property
Public Property Set MainWindow(ByVal vfMainWindow As Object)
    Set moMainWindow = vfMainWindow
End Property
Public Property Get MainWindow() As Object
    Set MainWindow = moMainWindow
End Property
Public Function GetDateTime() As Date
    GetDateTime = GetGMTDateTime("DATETIME")
End Function
Public Function GetDate() As Date
    GetDate = GetGMTDateTime("DATE")
End Function
Public Function GetTime() As Date
    GetTime = GetGMTDateTime("TIME")
End Function
Public Function GetTimeBias() As Long
    GetTimeBias = GetGMTBias(False)
End Function
Public Function GetTotalTimeBias() As Long
    GetTotalTimeBias = GetGMTBias(True)
End Function

Private Function GetGMTBias(ByVal bTotal As Boolean) As Long
    Dim tzInfo As TIME_ZONE_INFORMATION
    GetTimeZoneInformation tzInfo
    
    
    If bTotal Then
       GetGMTBias = tzInfo.Bias + tzInfo.DaylightBias
    Else
       GetGMTBias = tzInfo.Bias
    End If
    
End Function

Private Function GetGMTDateTime(ByVal sDateType As String) As Variant

    Dim SysTime     As SYSTEMTIME
    Dim GMTDateTime As Date
    
    Call GetSystemTime(SysTime)
    With SysTime
    Select Case sDateType
    Case "DATETIME"
        GetGMTDateTime = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond)
    Case "TIME"
        GetGMTDateTime = CDate(.wHour & ":" & .wMinute & ":" & .wSecond)
    Case "DATE"
        GetGMTDateTime = CDate(.wMonth & "/" & .wDay & "/" & .wYear)
    End Select
    End With
End Function

Public Function LocalToGmt(ByVal vdLocal As Date) As Date: LocalToGmt = DateConvert(vdLocal, 1): End Function
Public Function GmtToLocal(ByVal vdGmt As Date) As Date:   GmtToLocal = DateConvert(vdGmt, -1): End Function

Private Function DateConvert(ByVal vDate As Date, ByVal iDirection As Integer) As Date
    Dim tzInfo As TIME_ZONE_INFORMATION
    Dim TimeZone As Long
    
    'SCR 29070  11/12/02  Need to check the return code to
    'determine if Day Light Savings offset should be included in calculation.
    TimeZone = GetTimeZoneInformation(tzInfo)
                                        'SCR 65405
    If TimeZone = TIME_ZONE_ID_STANDARD Or TimeZone = TIME_ZONE_ID_UNKNOWN Then
        DateConvert = DateAdd("n", iDirection * (tzInfo.Bias), vDate)
    ElseIf TimeZone = TIME_ZONE_ID_DAYLIGHT Then
        DateConvert = DateAdd("n", iDirection * (tzInfo.Bias + tzInfo.DaylightBias), vDate)
    Else
        Err.Raise 1000, "Comm.Comm.DateConvert", "Can not convert Date.  Invalid Time Zone Information."
    End If

End Function

Public Function GetVisualIndicators() As Scripting.Dictionary
    Set GetVisualIndicators = mCWComm.GetVisualIndicators
End Function
Public Function InitializeOLEServer(NCID_Application As String) As Object
    Set InitializeOLEServer = mCWComm.InitializeOLEServer(NCID_Application)
End Function

Public Function SubscribeToMessage(ByVal SourceServerNCID As String, ByVal MessageType As Integer, ByVal DestinationServerNCID As String) As Boolean
    SubscribeToMessage = mCWComm.SubscribeToMessage(SourceServerNCID, MessageType, DestinationServerNCID)
End Function
Public Function UnSubscribeToMessage(ByVal SourceServerNCID As String, ByVal MessageType As Integer, ByVal DestinationServerNCID As String) As Boolean
    UnSubscribeToMessage = mCWComm.UnSubscribeToMessage(SourceServerNCID, MessageType, DestinationServerNCID)
End Function

Public Function RemoveListenerForMessage(ByVal SourceServerNCID As String, ByVal MessageType As Integer, ByVal DestinationServerNCID As String) As Boolean
    RemoveListenerForMessage = mCWComm.RemoveListenerForMessage(SourceServerNCID, MessageType, DestinationServerNCID)
End Function

Public Function GetSecurityObject() As String
    GetSecurityObject = mCWComm.GetSecurityObject()
End Function
Public Function GetMedcinIP() As String
    GetMedcinIP = mCWComm.GetMedcinIP()
End Function
Public Function GetMedcinPort() As String
    GetMedcinPort = mCWComm.GetMedcinPort()
End Function
Public Function GetMedcinState() As String
    GetMedcinState = mCWComm.GetMedcinState()
End Function
Public Function SetMedcinState(sNewState As String) As Boolean
    SetMedcinState = mCWComm.SetMedcinState(sNewState)
End Function

Public Property Get OE_MEDS_STATE() As Long
    OE_MEDS_STATE = mlOEMedsState
End Property
Public Property Let OE_MEDS_STATE(ByVal vNewValue As Long)
    mlOEMedsState = vNewValue
End Property

Public Function GetViewMenuItemAttributes() As Collection
    Set GetViewMenuItemAttributes = mCWComm.GetViewMenuItemAttributes
End Function
Public Function GetOLEServers() As Collection
    Set GetOLEServers = mCWComm.GetOLEServers
End Function
Public Function GetPatientIDLine() As String
    GetPatientIDLine = mCWComm.GetPatientIDLine
End Function


Public Function GetLocalTimeZoneABV() As String
    Dim sFTZ As String
    Dim sATZ As String
On Error GoTo errorHandler
    
    GetLocalTimeZone sFTZ, sATZ
    
    GetLocalTimeZoneABV = sATZ
    
    Exit Function
errorHandler:
    gShowVBError Err.Number, Err.Description, TypeName(Me) & ".GetLocalTimeZoneABV", vbCritical
End Function

'This function returns the strings representing the local time zone, i.e. "EST" and "Eastern Standard Time".
Public Function GetLocalTimeZone(FullTimeZoneString As String, AbvTimeZoneString As String) As Boolean
    Dim tzInfo As TIME_ZONE_INFORMATION
    Dim dStandard As Date
    Dim dDaylight As Date
    Dim dDate As Date
    Dim aName() As Integer
    Dim i As Integer
On Error GoTo errorHandler
    
    GetLocalTimeZone = True

    Call GetTimeZoneInformation(tzInfo)
    
    dDate = Now
    With tzInfo.StandardDate
        dStandard = DateSerial(VBA.Year(dDate), .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
    With tzInfo.DaylightDate
        dDaylight = DateSerial(VBA.Year(dDate), .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
    If dStandard > dDaylight Then
        If dDate >= dDaylight And dDate < dStandard Then
            aName = tzInfo.DaylightName
        Else
            aName = tzInfo.StandardName
        End If
    Else
        If dDate >= dStandard And dDate < dDaylight Then
            aName = tzInfo.StandardName
        Else
            aName = tzInfo.DaylightName
        End If
    End If
    
    FullTimeZoneString = Chr$(aName(0))
    AbvTimeZoneString = Chr$(aName(0))
    
    For i = 0 To UBound(aName)
        If aName(i) = 0 Then
            Exit For
        End If
        FullTimeZoneString = Chr$(aName(i))
        If Chr$(aName(i)) = " " Then
            If i < UBound(aName) Then
                AbvTimeZoneString = AbvTimeZoneString & Chr$(aName(i + 1))
            End If
        End If
    Next i
    
    Exit Function
errorHandler:
    gShowVBError Err.Number, Err.Description, TypeName(Me) & ".GetLocalTimeZone", vbCritical
    GetLocalTimeZone = False
End Function

'This function returns the date formatted with the time zone information on the end.
Public Function FormatDTG(ByVal DTG As Date) As String
  FormatDTG = Format$(DTG, gsDefaultDateTime) & " " & GetLocalTimeZoneABV()
End Function

