VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        LibTimer.cls - mclsLibTimer - Timer class
'''
'''Description:   Allows to track elapsed time
'''
'''
'''Suggested Use:
'''
'''Note:
'''
'''Modules Required: clsTimerEvent
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME      As String = "mclsLibTimer."

Private Enum mnum_TET_TimeEventType
   lTET_DATE
   lTET_DATE_ADD
   lTET_INTERVAL
End Enum

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mdStartSeconds           As Double
Private mtStartDate              As Date

Public pbLeadingZeroes           As Boolean
Public pbLongFormat              As Boolean

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private WithEvents mtmrEvents    As Timer
Attribute mtmrEvents.VB_VarHelpID = -1
Private mcolEvents               As Collection

Public Event pLibTimerEvent(ByVal tEventDateTime As Date, ByVal dDelay As Double, ByVal sMsg As String, ByVal sID As String, ByVal lCallNo As Long, ByVal vTag As Variant)



Private Sub M_ActivateTimer(ByVal bActivate As Boolean)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' bActivate        Whether to activate or diactivate timer
'''
'''Description:      Based on  request and current status
'''                  Increment or Decrement timer use count, which in turn
'''                  may activate or diactivate timer
'''
'''Note:
'''
'''Warning:          Do not send request to activate or diactivate
'''                  If change of status is not required - it will mess total timer use count
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "M_ActivateTimer"
   Static bActive_static      As Boolean
   
   '''If requested to activate and not already active
   If bActivate And (Not bActive_static) Then
      '''Activate timer - Increment timer use count
      Call Lib_TimerEnable(True)
      bActive_static = True
   '''If requested to diactivate and timer is active
   ElseIf Not bActivate And bActive_static Then
      '''Diactivate timer - Decrement timer use count
      Call Lib_TimerEnable(False)
      bActive_static = False
   End If

End Sub


Private Function M_bEventExists(ByVal sKey As String, ByVal bRemove As Boolean) As Boolean
'''Called by:
   Dim bResult    As Boolean
   Dim objEvent   As clsTimerEvent
   
On Error GoTo ErrHnd

   If Not mcolEvents Is Nothing Then
      For Each objEvent In mcolEvents
         If objEvent.psKey = sKey Then
            bResult = True
            If bRemove Then
               Call mcolEvents.Remove(sKey)
               bResult = False
            End If
         End If
      Next objEvent
   End If
   
   M_bEventExists = bResult
   
Exit Function
ErrHnd:
   M_bEventExists = False
   
End Function

Public Function P_bEventActivate(ByVal sID As String, ByVal bActivate As Boolean) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sID             ID of the specified event
''' bActivate        Indicates whether to activate or diactivate specified event
'''
'''Description:      Activate or diactivate specified event
'''
'''Note:
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     true
'''                  else
'''                     false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_bEventActivate"
   
   Dim objEvent               As New clsTimerEvent
   Dim bResult                As Boolean
      
On Error GoTo ErrHnd
   If Not mcolEvents Is Nothing Then
      For Each objEvent In mcolEvents
         '''If event corresponding to specified id is found
         If objEvent.psID = sID Then
            objEvent.pbActive = bActivate
            '''Note:
            '''objEvent.psID  is referencing collection key of the event, which is unique
            '''objEvent.psID is identifying event to the client and does not
            '''have to be unique - Therefore do not exit the loop because there
            '''can be more then one event with the same objEvent.psID
            bResult = True
         End If
      Next
   End If
   '''If event found and activated
   If bResult Then
      If bActivate Then
         '''Make sure that timer is turned on
         Call M_ActivateTimer(True)
      End If
   Else
      Call gobjLibApp.LibApp_Error _
          (1, gobjLibStr.LibStr_sBuildSentence(gsEVENT, gsIS, gsNOT, gsFOUND), _
           msMODULE_NAME, sPROC_NAME, gsID & " :" & sID _
          )
   End If
   
   P_bEventActivate = bResult
   
Exit Function
ErrHnd:
   P_bEventActivate = False
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsID & " :" & sID)

End Function


Public Function P_bEventsClearAll() As Boolean
'''************************************************************************
'''Called By:     Any
'''Arguments:
'''
'''Description:      Remove all events
'''
'''Note:
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     true
'''                  else
'''                     false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_bEventsClearAll"
   
On Error GoTo ErrHnd

   Set mcolEvents = Nothing
   
   P_bEventsClearAll = True
   
Exit Function
ErrHnd:
   P_bEventsClearAll = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function


Public Function P_bEventRemove(ByVal sID As String) As Boolean
'''************************************************************************
'''Called By:        P_bEventExists
'''
'''Arguments:
''' sID
''' ID of the Event requested to be removed
'''
'''Description:      Remove specified event
'''
'''Note:
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     true
'''                  else
'''                     false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_bEventRemove"
   
   Dim objEvent               As New clsTimerEvent
   Dim bResult                As Boolean
   Dim sKey                   As String
      
On Error GoTo ErrHnd

   If Not mcolEvents Is Nothing Then
      For Each objEvent In mcolEvents
         '''If event corresponding to specified id is found
         If objEvent.psID = sID Then
            sKey = objEvent.psKey
            Call mcolEvents.Remove(sKey)
            '''Note:
            '''objEvent.psKey is referencing collection key of the event, which is unique
            '''objEvent.psID is identifying event to the client and does not
            '''have to be unique - Therefore do not exit the loop because there
            '''can be more then one event with the same objEvent.psID
            bResult = True
         End If
      Next
   End If
   If Not bResult Then
      Call gobjLibApp.LibApp_Error _
          (1, gobjLibStr.LibStr_sBuildSentence(gsEVENT, gsIS, gsNOT, gsFOUND), _
           msMODULE_NAME, sPROC_NAME, gsKEY & " :" & sID _
          )
   End If
   
   P_bEventRemove = bResult
   
Exit Function
ErrHnd:
   P_bEventRemove = False
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsID & " :" & sID)

End Function

Public Function P_dElapsedTimeUnits _
      (ByVal sTimeUnit As String, ByVal tDateStart As Date, ByVal tDateEnd As Date) As Double
'''************************************************************************
'''Called By:           M_dElapsedTimeUnits
'''                     mtmrEvents_Timer or any
'''Arguments:
''' sTimeUnit           Determines time unit of return value
'''                        Compliant with DateDiff interval values
''' tDateStart          Start of the time period
''' tDateEnd            End of the time period
'''
'''Description:         Determine real number of requested time units
'''                     that passed between during specified time period
'''Note:
'''
'''Updates Globals:
'''Returns:             If succcessful
'''                        real number of requested time units
'''                     else
'''                        plERR_TIMER and error is raised
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_dElapsedTimeUnits"

   Dim tDateFromReminder   As Date
   Dim dResult             As Double
   Dim dSecondsWhole       As Double
   Dim dSecondsReminder    As Double
   Dim lPeriodUnits        As Long
   
On Error GoTo ErrHnd

   '''get whole number of units and reminder date range
   dResult = P_lDateDiffRoundDown(sTimeUnit, tDateStart, tDateEnd, tDateFromReminder)
   '''''' Determine  fractional part
   '''''' 1. Average seconds per time unit within period
   '''Get number of time units for the period rounded up to the nearest integer
   lPeriodUnits = DateDiff(sTimeUnit, tDateStart, tDateEnd)
   If lPeriodUnits = 0 Then
      '''Make sure we have at least on unit
      If tDateStart >= tDateEnd Then
         lPeriodUnits = -1
      Else
         lPeriodUnits = 1
      End If
   End If
   '''Determine total number of seconds for the rounded up period
   dSecondsWhole = DateDiff("s", tDateStart, DateAdd(sTimeUnit, lPeriodUnits, tDateStart))
   '''Determine denominator
   '''If more then one unit within period
   If Abs(lPeriodUnits) > 1 Then
      '''Average number of seconds in one time unit within period
      dSecondsWhole = dSecondsWhole / lPeriodUnits
   End If
   '''''' 2. Determine how many seconds in fractional part
   dSecondsReminder = P_lDateDiffRoundDown("s", tDateFromReminder, tDateEnd)

   '''If there is fractional part
   If Not dSecondsWhole = 0 Then
      '''Add fractional part - Abs to keep the sign -/-=+
      dResult = dResult + dSecondsReminder / Abs(dSecondsWhole)
   End If
   
   P_dElapsedTimeUnits = dResult
   
Exit Function
ErrHnd:
   P_dElapsedTimeUnits = plERR_TIMER
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sTimeUnit, True)

End Function

Public Function P_sAddEventRegularInterval _
      (ByVal sInterval As String, ByVal lNumber As Long, ByVal sMsg As String, _
       Optional ByVal sID As String, Optional ByVal vTag As Variant, _
       Optional bOverwrite As Boolean = True _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sInterval     Indicates time unit type specified in lNumber
'''                  Valid values are: d, h, n, s
''' lNumber       Specified number of time units that should pass between events
''' sMsg          Message requested to be returnd with event
''' sID          Optional event ID. Ignored if empty, returned instead of the generated ID
'''               if specified
''' vTag          Generic data requested to be associated with new event
''' bOverwrite
''' Indicates whether to replace all events with matching ID with this event
''' Ignored if sID is empty
'''
'''Description:   Initiate an event to be triggered when every specified number of units
'''
'''Note:          In order to receive events object of this class must be declared with events
'''               Negative number of units is enterpreted as an error
'''               Zero will be interpreted as:
'''               Fire as soon as possible - which currently will be in about 0.5 second
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     if Key value is not specified
'''                        New ID of the event
'''                     else
'''                        Key value
'''                  else
'''                     empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_sAddEventRegularInterval"

On Error GoTo ErrHnd

   sID = M_sAddEvent _
        (lTET_INTERVAL, P_lConvertTimeUnits(sInterval, lNumber, "s"), sMsg, sID, vTag, bOverwrite)

   P_sAddEventRegularInterval = sID
   
Exit Function
ErrHnd:
   P_sAddEventRegularInterval = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sMsg)

End Function


Public Function P_bEventExists _
(ByVal sID As String, Optional ByRef lCount As Long, Optional bRemove As Boolean) As Boolean
'''************************************************************************
'''Called By:     M_sAddEvent
'''
'''Arguments:
''' sID
'''   ID of the Event requested to be evaluated and optionaly returned
''' lCount
'''   Number of events with specified ID. Represents precondition count
''' bRemove
'''   Optional request to remove all mathing events
'''
'''Description:      Determine whether specified event exists
'''                  Optionally return number of events with matching ID
'''Note:
''' If remove is requested count may be returned positive even if function returns false
'''
'''Updates Globals:
'''Returns:          If specified event was found
'''                     true
'''                  else
'''                     false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_bEventExists"
   
   Dim bResult                As Boolean
   Dim objEvent               As clsTimerEvent
   
On Error GoTo ErrHnd

   lCount = 0
   If Not mcolEvents Is Nothing Then
      For Each objEvent In mcolEvents
         '''If event corresponding to specified id is found
         If objEvent.psID = sID Then
            bResult = True
            lCount = lCount + 1
            If bRemove Then
               Call P_bEventRemove(sID)
            End If
         End If
      Next objEvent
   End If
   
   P_bEventExists = bResult
   
Exit Function
ErrHnd:
   P_bEventExists = False
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsID & " :" & sID)

End Function

Public Property Get plERR_TIMER() As Long
   plERR_TIMER = -9999999
End Property


Public Property Get plSECONDS_IN_DAY() As Long
   plSECONDS_IN_DAY = 86400
End Property


Public Property Get plSECONDS_IN_HOUR() As Long
   plSECONDS_IN_HOUR = 3600
End Property

Private Function M_dElapsedTimeUnits(ByVal sTimeUnit As String) As Double
'''************************************************************************
'''Called By:           pdElapsedYears
'''                     pdElapsedMonth
'''                     pdElapsedWeeks
'''                     pdElapsedDays
'''                     pdElapsedHours
'''                     pdElapsedMinutes
'''
'''Arguments:
''' sTimeUnit           Determines time unit of return value
'''                        Compliant with DateDiff interval values
'''
'''Description:         Determine real number of requested time units
'''
'''Note:
'''
'''Updates Globals:
'''Returns:             If succcessful
'''                        real number of requested time units
'''                     else
'''                        plERR_TIMER and error is raised
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "M_dElapsedTimeUnits"

On Error GoTo ErrHnd

   M_dElapsedTimeUnits = P_dElapsedTimeUnits(sTimeUnit, ptStartTime, Now)

Exit Function
ErrHnd:
   M_dElapsedTimeUnits = plERR_TIMER
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sTimeUnit, True)

End Function

Private Function M_sAddEvent _
      (lEventType As mnum_TET_TimeEventType, _
       ByVal vTime As Variant, ByVal sMsg As String, ByVal sID As String, ByVal vTag As Variant, _
       bOverwrite As Boolean _
      ) As String
'''************************************************************************
'''Called By:     P_sAddEventDate
'''               P_sAddEventDateAdd
'''               P_sAddEventRegularInterval
'''
'''Arguments:
''' lEventType    Indicates how time of the requested event will be determined and how to interpred
'''               specified vTime:
'''                  specific date, elapsed seconds since start time, regular time interval in seconds
''' vTime         Based on type:
'''                  specific date, elapsed seconds since start time, time interval in seconds
'''
''' sMsg          Message requested to be returned with event
''' sID           ID requested to be associated with new event
'''               Ignored if empty, returned instead of the generated ID - if specified
''' vTag          Generic data requested to be associated with new event
''' bOverwrite
''' Indicates whether to replace all events with matching ID with this event
''' Ignored if sID is empty
'''
'''Description:   Initiate requested event
'''
'''Note:             sID value will not be trimed or modified in any way -
'''                  so if it's 3 spaces that's what will be returned
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     if sID is not specified
'''                        New ID of the event
'''                     else
'''                        sID value
'''                  else
'''                     empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "M_sAddEvent"
   
   Static lEventID_static     As Long
   Dim sEventKey              As String
   Dim sErrMsg                As String
   Dim objEvent               As New clsTimerEvent
   
On Error GoTo ErrHnd

   If bOverwrite Then
      If Len(sID) > 0 Then
         Call P_bEventExists(sID, , True)
      End If
   End If

   '''Based on event time set type and value of the event time
   Select Case lEventType
      Case lTET_DATE
         '''Make sure its a date type
         vTime = CDate(vTime)
      Case lTET_DATE_ADD
         '''Determine time of event by adding specified number of seconds
         vTime = DateAdd("s", vTime, Now)
      Case lTET_INTERVAL
         '''Make sure its a long type
         vTime = CLng(vTime)
      Case Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsTYPE, lEventType, vbCrLf, sMsg)
         Call gobjLibApp.LibApp_Error _
            (1, gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsTYPE), msMODULE_NAME, sPROC_NAME, sErrMsg)
   End Select
   '''If no errors
   If sErrMsg = "" Then
      '''Initiate the event
      '''If first call
      If mcolEvents Is Nothing Then
         Set mcolEvents = New Collection
      End If
      '''Generate new KEY
      lEventID_static = lEventID_static + 1
      sEventKey = "E" & CStr(lEventID_static)
      '''If Key is not specified by the caller
      If sID = "" Then
         '''Return newly generated key
         sID = sEventKey
      End If
   
      With objEvent
         .psKey = sEventKey
         .pvTime = vTime
         .psMsg = sMsg
         .psID = sID
         .ptTimeStart = Now
         .pvTag = vTag
         .pbActive = True
      End With
      
      Call mcolEvents.Add(objEvent, sEventKey)
      Set objEvent = Nothing
      Call M_ActivateTimer(True)
   End If
   
   M_sAddEvent = sID
   
Exit Function
ErrHnd:
   M_sAddEvent = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sMsg)

End Function

Public Function P_lDateDiffRoundDown _
      (ByVal sTimeUnit As String, ByVal tDateFrom As Date, ByVal tDateTo As Date, _
       Optional ByRef tAdjDateFrom As Date _
      ) As Long
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' tDateFrom           Start date/time of the interval
''' tDateTo             End date/time of the interval
''' sTimeUnit           Determines time unit
''' tAdjDateFrom        Optional Adjusted value of the date from representing
'''                        remaining date range
'''
'''Description:         Determine number of requested time units within specified
'''                     Date range.
'''
'''Note:                DateDiff always rounds e.g. 1.1 months counted as 2
'''                     but we want to count it as 1 and return reminder by adjusting
'''                     date from
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     Rounded down number of requested time within specified Date range
'''                  else
'''                     plERR_TIMER and error is raised
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_lDateDiffRoundDown"
   
   Dim lUnits                 As Long
   Dim tSwap                  As Date
   
On Error GoTo ErrHnd

   lUnits = DateDiff(sTimeUnit, tDateFrom, tDateTo)
   If lUnits > 0 Then
      If DateAdd(sTimeUnit, lUnits, tDateFrom) > tDateTo Then
         lUnits = lUnits - 1
      End If
   ElseIf lUnits < 0 Then
      If DateAdd(sTimeUnit, lUnits, tDateFrom) < tDateTo Then
         lUnits = lUnits + 1
      End If
   End If
   
   tAdjDateFrom = DateAdd(sTimeUnit, lUnits, tDateFrom)
   
   P_lDateDiffRoundDown = lUnits
   
Exit Function
ErrHnd:
   P_lDateDiffRoundDown = plERR_TIMER
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sTimeUnit, True)
   
End Function
Public Sub P_SetStartTime(Optional ByVal tDateTime As Date)
   '''If date is not specified
   If tDateTime = 0 Then
      '''Default to the current date time
      mtStartDate = Date
      mdStartSeconds = Timer
   Else
      mtStartDate = CDate(Format$(tDateTime, "mm/dd/yyyy"))
      '''Set to the number of second since the beginning of the specified date
      mdStartSeconds = DateDiff("s", mtStartDate, tDateTime)
   End If
   
End Sub

Public Function P_lConvertTimeUnits _
      (ByVal sIntervalFrom As String, ByVal lNumberFrom As Long, ByVal sIntervalTo As String) As Long
'''************************************************************************
'''Called By:     P_sAddEventDateAdd or Any
'''
'''Arguments:
''' sIntervalFrom Indicates initial time unit type
'''                  Valid values are: d, h, n, s
''' lNumberFrom   Initial number of time units - interpreted based on sIntervalFrom
''' sIntervalTo   Time unit type requested for conversion
'''
'''Description:   Convert specified number of units of the sIntervalFrom time unit type
'''               into number of units of the sIntervalTo time unit type
'''Note:
'''
'''Updates Globals:
'''Returns:             If succcessful
'''                        Converted number of time units
'''                     else
'''                        plERR_TIMER and error is raised
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_lConvertTimeUnits"

   Dim lResult                As Long
   Dim lSeconds               As Long
   Dim bPopErr                As Boolean
   
On Error GoTo ErrHnd

   lResult = plERR_TIMER
   bPopErr = True
   '''First convert to seconds
   Select Case sIntervalFrom
      Case "s"
         lSeconds = lNumberFrom
      Case "n"
         lSeconds = lNumberFrom * 60
      Case "h"
         lSeconds = lNumberFrom * plSECONDS_IN_HOUR
      Case "d"
         lSeconds = lNumberFrom * plSECONDS_IN_DAY
      Case Else
         bPopErr = False '''Cancel display and log error options in error trap
         Call gobjLibApp.LibApp_Error _
         (1, gobjLibStr.LibStr_sBuildSentence _
                       (gsFROM, "-", gsSPECIFIED, gsUNIT, gsTYPE, gsIS, gsNOT, gsSUPPORTED), _
          msMODULE_NAME, sPROC_NAME, _
          gobjLibStr.LibStr_sBuildSentence(gsTYPE, sIntervalFrom), True _
         )
   End Select
   '''Convert resulted seconds to requested time units
   Select Case sIntervalTo
      Case "s"
         lResult = lSeconds
      Case "n"
         lResult = lSeconds / 60
      Case "h"
         lResult = lSeconds / plSECONDS_IN_HOUR
      Case "d"
         lResult = lSeconds / plSECONDS_IN_DAY
      Case Else
         bPopErr = False '''Cancel display and log error options in error trap
         Call gobjLibApp.LibApp_Error _
         (1, gobjLibStr.LibStr_sBuildSentence _
                        (gsTO, "-", gsSPECIFIED, gsUNIT, gsTYPE, gsIS, gsNOT, gsSUPPORTED), _
          msMODULE_NAME, sPROC_NAME, _
          gobjLibStr.LibStr_sBuildSentence(gsTYPE, sIntervalTo), True _
         )
   End Select
   
   P_lConvertTimeUnits = lResult
   
Exit Function
ErrHnd:
   P_lConvertTimeUnits = plERR_TIMER
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True, bPopErr, bPopErr)

End Function

Public Function P_sAddEventDate _
      (ByVal tDateTime As Date, ByVal sMsg As String, _
       Optional ByVal sID As String, Optional ByVal vTag As Variant, _
       Optional bOverwrite As Boolean = True _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' tDateTime     date of the requested event
''' sMsg          Message requested to be returnd with event
''' sID          Optional event ID. Ignored if empty, returned instead of the generated ID
'''               if specified
''' vTag          Generic data requested to be associated with new event
''' bOverwrite
''' Indicates whether to replace all events with matching ID with this event
''' Ignored if sID is empty
'''
'''Description:   Initiate an event to be triggered on specified date/time
'''
'''Note:             In order to receive events object of this class must be declared with events.
'''                  This event will occur only ones
'''
'''Warning:
''' if bOverwrite is true
'''   all events with matching ID will be removed priour to adding new one
'''
'''Returns:          If succcessful
'''                     if sID is not specified
'''                        New ID of the event
'''                     else
'''                        sID value
'''                  else
'''                     empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_sAddEventDate"
   
On Error GoTo ErrHnd

   P_sAddEventDate = M_sAddEvent(lTET_DATE, tDateTime, sMsg, sID, vTag, bOverwrite)
   
Exit Function
ErrHnd:
   P_sAddEventDate = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sMsg)

End Function

Public Function P_sAddEventDateAdd _
      (ByVal sInterval As String, ByVal lNumber As Long, ByVal sMsg As String, _
       Optional ByVal sID As String, Optional ByVal vTag As Variant, _
       Optional bOverwrite As Boolean = True _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sInterval     Indicates time unit type specified in lNumber
'''                  Valid values are: d, h, n, s
''' lNumber       Specifid number of units since now
''' sMsg          Message requested to be returnd with event
''' sID          Optional event ID. Ignored if empty, returned instead of the generated ID
'''               if specified
''' vTag          Generic data requested to be associated with new event
''' bOverwrite
''' Indicates whether to replace all events with matching ID with this event
''' Ignored if sID is empty
'''
'''Description:   Initiate an event to be triggered when specified number of units will pass
'''               after this event was initiate (now)
'''
'''Note:          In order to receive events object of this class must be declared with events
'''               If number of units is negative it will be enterpreted as zero
'''               This event will occur only ones
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     if sID is not specified
'''                        New ID of the event
'''                     else
'''                        sID value
'''                  else
'''                     empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_sAddEventDateAdd"

On Error GoTo ErrHnd

   P_sAddEventDateAdd = M_sAddEvent _
               (lTET_DATE_ADD, P_lConvertTimeUnits(sInterval, lNumber, "s"), sMsg, sID, vTag, bOverwrite)
   
Exit Function
ErrHnd:
   P_sAddEventDateAdd = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sMsg)

End Function

Public Property Get pdElapsedDays() As Double
   pdElapsedDays = M_dElapsedTimeUnits("d")
   
End Property

Public Property Get pdElapsedHours() As Double
   pdElapsedHours = M_dElapsedTimeUnits("h")
End Property

Public Property Get pdElapsedMinutes() As Double
   pdElapsedMinutes = M_dElapsedTimeUnits("n")
End Property

Public Property Get pdElapsedMonths() As Double
   pdElapsedMonths = M_dElapsedTimeUnits("m")
End Property

Public Property Get pdElapsedSeconds() As Double
'''************************************************************************
'''Called By:           None
'''
'''Arguments:
'''
'''Description:         Determine and return number of seconds elapsed since
'''                     start time including fractional part
'''Note:
'''
'''Updates Globals:
'''Returns:             If succcessful
'''                        real number of seconds since start time
'''                     else
'''                        plERR_TIMER and error is raised
'''
'''Programmer:          Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "pdElapsedSeconds"
   Dim dTimer                 As Double
   
On Error GoTo ErrHnd

   dTimer = Timer '''To avoid rounding errors - make sure that double is used for evaluation
   pdElapsedSeconds = DateDiff("s", mtStartDate, Date) + dTimer - Abs(mdStartSeconds)
   
Exit Property
ErrHnd:
   pdElapsedSeconds = plERR_TIMER
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Get pdElapsedWeeks() As Double
   pdElapsedWeeks = M_dElapsedTimeUnits("ww")
End Property

Public Property Get pdElapsedYears() As Double
   pdElapsedYears = M_dElapsedTimeUnits("yyyy")
End Property

Public Property Get ptStartTime() As Date
   ptStartTime = DateAdd("s", mdStartSeconds, mtStartDate)
End Property


Public Property Get psElapsedTime() As String
   psElapsedTime = P_sTimeInterval(ptStartTime, Now, pbLeadingZeroes, , , pbLongFormat)
End Property
Public Function P_sTimeInterval _
      (ByVal tDateFrom As Date, ByVal tDateTo As Date, _
       Optional bLeadingZeroes As Boolean = False, _
       Optional sMaxTimeUnit As String = "y", _
       Optional lMaxUnitLen As Long = 0, _
       Optional bLongFormat As Boolean = True _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' tDateFrom           Start date/time of the interval
''' tDateTo             End date/time of the interval
''' lTotalSeconds       Number of seconds specified
''' bLeadingZeroes      Indicates whether to return leading zeroes
''' sMaxTimeUnit        Determines greatest time unit of the returned time interval
'''                        Valid values are: y, m, w, d, h, n, s
'''                        e.g.  "h" means that greatest time unit will be hours
'''                              so 1 day 1 hour will be expressed as 25 hours
'''                        defaults to h - hours
''' lMaxUnitLen         Indicates static number of characters for max unit
'''                     If specified number of characters for lesser units will be 2
'''                        Dynamic if Zero
''' bLongFormat         Indicates whether returned time format
'''                     Defaults to false - [yy:mm:dd:hh:nn:]ss
'''                        if true will return [y years m months d days ...] ss seconds
'''
'''Description:         Based on specified Max time unit and format
'''                     Convert specified number of seconds into time interval expression
'''
'''Note:                This function returns absolute value of the interval - time between
'''                     DateFrom and DateTo regardless of whether DateFrom is before or after
'''                     the DateTo
'''Updates Globals:
'''Returns:          If succcessful
'''                     specified number of seconds in [hh:]mm:ss format
'''                  else
'''                     Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "P_sTimeInterval"
   Const lNO_OF_UNITS         As Long = 7
   
   Dim sResult                As String
   Dim lYears                 As Long
   Dim lMonths                As Long
   Dim lWeeks                 As Long
   Dim lDays                  As Long
   Dim lHours                 As Long
   Dim lMinutes               As Long
   Dim lSeconds               As Long
   Dim lTimeUnit              As Long
   Dim lMaxTimeUnit           As Long
   Dim lUnit                  As Long
   Dim sFormat                As String
   Dim tSwap                  As Date
   
On Error GoTo ErrHnd

   If tDateFrom > tDateTo Then
      tSwap = tDateFrom
      tDateFrom = tDateTo
      tDateTo = tSwap
   End If
   '''Determine starting point
   lMaxTimeUnit = InStr("ymwdhns", Left$(sMaxTimeUnit, 1))
   '''If specified and found
   If lMaxTimeUnit > 0 And sMaxTimeUnit <> "" Then
      For lTimeUnit = lMaxTimeUnit To lNO_OF_UNITS
         Select Case lTimeUnit
            Case 1   '''Years
               lYears = P_lDateDiffRoundDown("yyyy", tDateFrom, tDateTo, tDateFrom)
            Case 2   '''Months
               lMonths = P_lDateDiffRoundDown("m", tDateFrom, tDateTo, tDateFrom)
            Case 3   '''Weeks
               lWeeks = P_lDateDiffRoundDown("ww", tDateFrom, tDateTo, tDateFrom)
            Case 4   '''Days
               lDays = P_lDateDiffRoundDown("d", tDateFrom, tDateTo, tDateFrom)
            Case 5   '''Hours
               lHours = P_lDateDiffRoundDown("h", tDateFrom, tDateTo, tDateFrom)
            Case 6   '''Minutes
               lMinutes = P_lDateDiffRoundDown("n", tDateFrom, tDateTo, tDateFrom)
            Case 7   '''Seconds
               lSeconds = DateDiff("s", tDateFrom, tDateTo)
         End Select
      Next lTimeUnit
      '''Format output
      If lMaxUnitLen > 0 Then
         sFormat = String(lMaxUnitLen, "0")
      End If
      For lTimeUnit = lMaxTimeUnit To lNO_OF_UNITS
         lUnit = Choose(lTimeUnit, lYears, lMonths, lWeeks, lDays, lHours, lMinutes, lSeconds)
         If bLeadingZeroes Or lUnit > 0 Or Not (sResult = "") Then
            If bLongFormat Then
               sResult = sResult & Format$(lUnit, sFormat) & " " _
               & Choose(lTimeUnit, "years", "months", "weeks", "days", "hours", "minutes", "seconds") _
               & " "
            Else
               sResult = sResult & ":" & Format$(lUnit, sFormat)
            End If
         End If
         If lMaxUnitLen > 0 Then
            '''Set format for lesser units
            sFormat = "00"
         End If
      Next lTimeUnit
      
      If Not bLongFormat Then
         sResult = Mid$(sResult, 2)
      End If
   Else
      Call gobjLibApp.LibApp_Error _
           (1, gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsTIME, gsUNIT, gsSPECIFIED), _
            msMODULE_NAME, sPROC_NAME, sMaxTimeUnit _
           )
   End If
   
   P_sTimeInterval = sResult
   
Exit Function
ErrHnd:
   P_sTimeInterval = ""
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsCASE & lTimeUnit)
End Function

Private Sub Class_Initialize()
   pbLongFormat = True
   '''Default to the time when object was created
   Call P_SetStartTime
   Set mtmrEvents = Lib_tmr

End Sub
Private Sub Class_Terminate()
On Error Resume Next
   Set mcolEvents = Nothing
   Set mtmrEvents = Nothing
   
End Sub


Private Sub mtmrEvents_Timer()
'''************************************************************************
'''Called By:           None
'''
'''Arguments:
'''
'''Description:         If event is requested
'''                        Rise it
'''
'''Note:
'''pLibTimerEvent arguments:
''' tEventDateTime Expected time of the event
''' dDelay Indicates delay time in seconds between time when event was suppose to occur
'''        and time when it actually occurred
''' sMsg Text associated with the event when it was created
''' sID ID of the Event associated with the event when it was created
''' lCallNo Indicates how many times event was raised
''' vTag Optional data tag associated with the event when it was created
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:          Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "mtmrEvents_Timer"
   
   Dim colEvent               As Collection
   Dim sEventKey              As String
   Dim lActive_Count          As Long
   Dim objEvent               As clsTimerEvent
   Dim dDelay                 As Double
   Dim tExpEvnTime            As Date
   
On Error GoTo ErrHnd

      '''If there is anything to process
100   If Not mcolEvents Is Nothing Then
200      For Each objEvent In mcolEvents
300         With objEvent
               '''if event is active
400            If .pbActive Then
420               lActive_Count = lActive_Count + 1
                  '''If time is specific time
500               If VarType(.pvTime) = vbDate Then
520                  dDelay = P_dElapsedTimeUnits("s", .pvTime, Now)
                     '''If specified time of the event was reached
600                  If dDelay >= 0 Then
620                     .plCallNo = .plCallNo + 1 '''count number of times event was called
640                     RaiseEvent pLibTimerEvent(.pvTime, dDelay, .psMsg, .psID, .plCallNo, .pvTag)
660                     sEventKey = .psKey
                        '''Specific time should occure only ones
                        '''Remove this event if caller did not remove this event when it was raised above
800                     Call M_bEventExists(sEventKey, True)
                     End If
                  Else
                     '''Event should occure at the regular time intervals in seconds
1000                 tExpEvnTime = DateAdd("s", .pvTime, .ptTimeStart)
1020                 dDelay = P_dElapsedTimeUnits("s", tExpEvnTime, Now)
1100                 If dDelay >= 0 Then
1120                    .plCallNo = .plCallNo + 1 '''count number of times event was called
                        '''Reset start time for the interval
1140                    .ptTimeStart = Now
1160                    RaiseEvent pLibTimerEvent(tExpEvnTime, dDelay, .psMsg, .psID, .plCallNo, .pvTag)
                     End If
                  End If
               End If
            End With
         Next objEvent
      End If
      '''If not events need to be fired
2000  If lActive_Count < 1 Then
         '''Wait until at least one active event
2100     Call M_ActivateTimer(False)
      End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Sub


