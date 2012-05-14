VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xDebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************
'Class Name: cDebugLog
'
'Description: This class logs statistical information about each procedure
'             in a module in an application.  Each procedure in application
'             must call the debuglog procedure once at the beginning and
'             once at the end in order for the statistical information to
'             be useful.
'Properties:
'
'Methods:
'   DebugStart      - Starts the debugging process
'   DebugStop       - Terminates the debugging process
'   DebugLog        - Writes Log information to a file and/or to the Screen.
'   DebugLogPing    - Writes Log information to a file and/or to the Screen but
'                     a system ping can be specified and recorded also.
'***************************************************************************
Option Explicit

'This record should be declared in each procedure and tracks the time elapsed
'while in a procedure
Rem X Public Type DebugRec
Rem X    TimerTime As Single
Rem X End Type

'Local Ini file that configuration information is stored
Private Const PGMINI = "DebugLog.ini"

'Section name where configuration information is stored under in the ini file (and OSH)
Private Const APPNAME = "Debug"

'This value is used to track the total time the application has been running
Private DebugTimer As Single


Private Sub DebugSaveGridData()

On Error GoTo ErrHandler

    Dim i%
    Dim Width As String
    
    For i% = 0 To giTOTAL_COLS - 1
        frmLog.fgLog.Row = 0
        frmLog.fgLog.Col = i%
        Rem winWritePrivateIni2 "fglog", "Left" & CStr(i%), (frmLog.fgLog.CellLeft), PGMINI
        Width = CStr(frmLog.fgLog.CellWidth)
        winWritePrivateIni3 "fglog", "Width" & CStr(i%), Width, gsDebugLogPath
    Next
    Exit Sub
ErrHandler:
    MsgBox "(DebugSaveGridData) Error:" & Error(Err)
    Exit Sub
End Sub
Private Sub DebugRestoreGridData()

On Error GoTo ErrHandler

    Dim i%
    Dim Left As Long
    Dim Width As Long
    On Error Resume Next
    For i% = 0 To giTOTAL_COLS - 1
        frmLog.fgLog.Row = 0
        Rem Left = CLng(winGetPrivateIni2$("fglog", "Left" & CStr(i%), PGMINI))
        Rem If Err <> 0 Then Exit For
        Width = CLng(winGetPrivateIni2$("fglog", "Width" & CStr(i%), PGMINI))
        If Err <> 0 Then Exit For
        frmLog.fgLog.Col = i%
        'frmLog.fgLog.ColLeft = Left
        frmLog.fgLog.ColWidth(i%) = Width
        If Err <> 0 Then MsgBox "Restore Error"
    Next
    Exit Sub
ErrHandler:
    MsgBox "(DebugRestoreGridData) Error:" & Error(Err)
    Exit Sub
End Sub
Private Sub DebugGetTheMode(sExe As String, sModule As String, sProc As String, tTheDebugMode As DebugModes)
'**********************************************************************************************************
'Sub Name: DebugGetTheMode
'
'Description: Determines the debug mode for a specific procedure
'
'Parameters:
'sExe - Name of calling executable or DLL
'sModule - Name of calling VB module (FORM or BAS module)
'sProc - Name of calling procedure
'ssOper - Name of operation being performed within procedure (i.e Begin, End, ..)
'tDR - Record that tracks the processing time spent in a module
'VarList - unlimited length list of values that contains the calling procedures parameter names and values
'           and variable names and values.  The parameters are specified first and a bar ("|") character
'           is used to separate the variables.  i.e. "p1", p1, "p2", p2, "|", "v1", v2
'**********************************************************************************************************
On Error GoTo ErrHandler

'Declarations ****************************************
Dim i As Integer 'array index counter

'Loop through the array (paModuleDebugModeArr) and check to see if the given procedure (Module, ect..) has
'a specific debug mode.
    tTheDebugMode = None
    For i = LBound(paModuleDebugModeArr) To UBound(paModuleDebugModeArr)
        If UCase(sExe) = UCase(paModuleDebugModeArr(i).Appl) Then
            If paModuleDebugModeArr(i).ModType = tApplication Then
                tTheDebugMode = paModuleDebugModeArr(i).DMode
            ElseIf UCase(sModule) = UCase(paModuleDebugModeArr(i).Modu) Then
                If paModuleDebugModeArr(i).ModType = tModule Then
                    tTheDebugMode = paModuleDebugModeArr(i).DMode
                ElseIf UCase(sProc) = UCase(paModuleDebugModeArr(i).Proc) Then
                    tTheDebugMode = paModuleDebugModeArr(i).DMode
                End If
            End If
        End If
    Next
Exit Sub

ErrHandler:
    MsgBox "(DebugGetTheMode) Error: " & Error(Err)
    Exit Sub
End Sub
Private Sub debugUpdateLog(msg As String)
' Purpose:
'
' Inputs:   Msg      - Message to log.
'
    On Error GoTo ErrHandler
    
    Dim LogFileNum As Long                  'dimmed
          
    LogFileNum = FreeFile
    Open gsDebugLogPath For Append Shared As LogFileNum
    Print #LogFileNum, msg
    Close LogFileNum
    
    Exit Sub

' An Error Occurred
ErrHandler:
    Dim ErrMsg$
    ErrMsg$ = Error$ & " - Writing to Debug Log Failed! "
    ErrMsg$ = ErrMsg$ & "Writing '" & gsDebugLogPath & "'"

    MsgBox ErrMsg$, vbCritical
    Exit Sub
End Sub


Private Sub debugLogPathSet(sLogPath As String)
    gsDebugLogPath = sLogPath

End Sub
Private Sub debugModeDefaultSet(DMode As DebugModes)
    
End Sub
Private Sub debugModeModuleSet(Module As String, DMode As DebugModes)

End Sub
Public Sub debugStart()
'*********************************************************************************
'Sub Name: debuStart
'
'Description: Initializes and starts debugging
'
'*********************************************************************************

On Error GoTo ErrHandler

Dim i As Integer 'Counter
Dim temp$
Dim MRec As ModRec
Dim IRec As IPRec
Dim Start%

'Get Control Mode (either ini file or OSH)

    On Error Resume Next
    
'******************************************************************
    gsIniPath = GetSupportPath("INI") & "\" & PGMINI
    giDefaultDebugMode = CInt(winGetPrivateIni3(APPNAME, "DefaultDebugMode", gsIniPath))
    If Err <> 0 Then
        MsgBox "Debug Error: invalid Default Debug Mode"
        Err = 0
    End If
    If giDefaultDebugMode = 3 Then Exit Sub
        
    
    giControlMode = CInt(winGetPrivateIni3(APPNAME, "ControlMode", gsIniPath))
    If Err <> 0 Then
        MsgBox "Debug Error: invalid Control Mode", vbCritical
        giControlMode = LocalControl
        Err = 0
    End If
    
    '<<< SCR 22541  MSolano 3/29/2002
    giARMOverride = CInt(winGetPrivateIni3(APPNAME, "ARMOverride", gsIniPath))
    If Err <> 0 Then
        giARMOverride = 0
        Err = 0
    End If
    '>>> SCR 22541
    
    
'Get settings from INI file
    'If giControlMode = LocalControl Then
        
'******************************************************************
        If giControlMode = CentralControl Then
            giDisplayMode = CInt(oshReadIni("DisplayMode"))
        Else
            giDisplayMode = CInt(winGetPrivateIni3(APPNAME, "DisplayMode", gsIniPath))
        End If
        If Err <> 0 Then
            MsgBox "Debug Error: invalid Display Mode"
            Err = 0
        End If

'******************************************************************
        If giControlMode = CentralControl Then
            gsDebugLogPath = oshReadIni("DebugLogPath")
        Else
            gsDebugLogPath = winGetPrivateIni3(APPNAME, "DebugLogPath", gsIniPath)
        End If
        If Err <> 0 Then
            MsgBox "Debug Error: invalid Default log file path"
            Err = 0
        End If
        
        gsCopyPath = winGetPrivateIni3(APPNAME, "CopyPath", gsIniPath)
        If Err <> 0 Then
            MsgBox "Debug Error: invalid Copy path"
            Err = 0
        End If
        
'******************************************************************
        If giControlMode = CentralControl Then
            giLogMove = CInt(oshReadIni("LogMove"))
        Else
            giLogMove = CInt(winGetPrivateIni3(APPNAME, "LogMove", gsIniPath))
        End If
        If Err <> 0 Then
            MsgBox "Debug Error: invalid Log Move value"
            Err = 0
        End If
        
'******************************************************************
        i% = 0
        Erase paModuleDebugModeArr
        giModuleDebugModeCount = 0
        Do
            On Error GoTo Skip
            i% = i% + 1
            Start% = 1
            If giControlMode = CentralControl Then
                temp$ = oshReadIni("Module" & CStr(i%))
            Else
                temp$ = winGetPrivateIni3(APPNAME, "Module" & CStr(i%), gsIniPath)
            End If
            If temp$ = "" Then Exit Do
                MRec.Appl = genStrTok$(Start%, temp$, ",")
                MRec.Modu = genStrTok$(Start%, temp$, ",")
                MRec.Proc = genStrTok$(Start%, temp$, ",")
                MRec.DMode = CInt(genStrTok$(Start, temp$, ","))
                If MRec.Modu = "" Then
                    MRec.ModType = tApplication
                ElseIf MRec.Proc = "" Then
                    MRec.ModType = tModule
                Else
                    MRec.ModType = tProcedure
                End If
            ReDim Preserve paModuleDebugModeArr(1 To i%)
            paModuleDebugModeArr(i%) = MRec
            giModuleDebugModeCount = giModuleDebugModeCount + 1
Skip:
        Loop
    
'******************************************************************
        i = 0
        Erase paIPArr
        piIPCount = 0
        Do
            On Error GoTo Skip2
            i = i + 1
            Start% = 1
            If giControlMode = CentralControl Then
                temp$ = oshReadIni("IP" & CStr(i))
            Else
                temp$ = winGetPrivateIni3(APPNAME, "IP" & CStr(i), gsIniPath)
            End If
            If temp$ = "" Then Exit Do
                IRec.Name = genStrTok$(Start%, temp$, ",")
                IRec.IP = genStrTok$(Start%, temp$, ",")
            ReDim Preserve paIPArr(1 To i)
            paIPArr(i) = IRec
            piIPCount = piIPCount + 1
Skip2:
        Loop

        If giDisplayMode = ToScreen Or giDisplayMode = Both Then
            frmLog.Show
            winPositionForm2 frmLog, gsDebugLogPath
            DebugRestoreGridData
            piLOGGING_FRM_LOADED = True
        End If
        If giDisplayMode = Both Or giDisplayMode = LogFile Then
            'genUpdateLog gsDebugLogPath, Now & "  Begin Logging *****************************************"
        End If


   

'Save Start Time of Application
    DebugTimer = Timer
    
    Exit Sub
    
ErrHandler:
    MsgBox "(debugStart) - Error:" & Error(Err)
    Exit Sub
End Sub
Public Sub debugStop()

On Error GoTo ErrHandler

    If giDisplayMode = Both Or giDisplayMode = ToScreen Then
        If piLOGGING_FRM_LOADED Then
            winSaveFormPosition2 frmLog, gsDebugLogPath
            DebugSaveGridData
            Unload frmLog
        End If
    End If
    If giDisplayMode = Both Or giDisplayMode = LogFile Then
        'genUpdateLog gsDebugLogPath, Now & "  Stop Logging *****************************************"
    End If
    If giLogMove = 1 And giDisplayMode <> ToScreen Then
        FileCopy gsDebugLogPath, gsCopyPath
        Kill gsDebugLogPath
    End If
    Exit Sub

ErrHandler:
    MsgBox "(debugStop) Error: " & Error(Err)
    Exit Sub
    
End Sub
Private Sub DebugGetBarIndex(iBarIndex As Integer, bBarExists As Boolean, vntVarList() As Variant)
'**********************************************************************************************************
'Sub Name: DebugGetBarIndex
'
'Description: Returns the locations of the Bar "|" character in the array
'
'Parameters:
'iBarIndex - Location in the vntVarList where the bar ("|") character is located
'bBarExists - True if the bar character was found, otherwise false
'vntVarList - array that contains a calling procedures variables (and parameters) names and values
'           and variable names and values.  The parameters are specified first and a bar ("|") character
'           is used to separate the variables.  i.e. "p1", p1, "p2", p2, "|", "v1", v2
'**********************************************************************************************************
On Error GoTo ErrHandler

'Declarations ****************************************
Dim i 'array index counter


'Find the bar character in the array
    iBarIndex = -1
    bBarExists = False
    On Error Resume Next
    For i = LBound(vntVarList) To UBound(vntVarList)
        If Err <> 0 Then
            iBarIndex = 0
            Exit Sub
        End If
        If TypeName(vntVarList(i)) = "String" Then
            If vntVarList(i) = "|" Then
                bBarExists = True
                iBarIndex = i
                Exit Sub
            End If
        End If
    Next
    iBarIndex = UBound(vntVarList) + 1
    Exit Sub
    
ErrHandler:
    MsgBox "(DebugGetBarIndex) Error: " & Error(Err)
    Exit Sub
End Sub
Private Sub DebugBuildParameterList(iBarIndex As Integer, vntVarList() As Variant, sParamList As String, iIsError As Integer)
'**********************************************************************************************************
'Sub Name: DebugBuildParameterList
'
'Description: Builds a bar separated listed of parameters passed in as a variant array.
'             i.e.  vntVarList = ("a",5,"b","abc","|","q",10) becomes
'                                "a='5' | b = 'abc'"
'             Notice the items after the "|" are ignored because those values are variables not parameters
'
'Parameters:
'iBarIndex - Location in the vntVarList where the bar ("|") character is located
'vntVarList - array that contains a calling procedures parameter names and values
'           and variable names and values.  The parameters are specified first and a bar ("|") character
'           is used to separate the variables.  i.e. "p1", p1, "p2", p2, "|", "v1", v2
'sParamList - a string with the list of parameters i.e. "a='5' | b = 'abc'"
'iIsError - Indicates whether there was an error.  Typically this would occur if there was an odd
'           number of values.
'**********************************************************************************************************
    On Error GoTo ErrHandler
    
'Declarations ****************************************
    Dim i As Integer 'Index for looping through the VntVarList array
    Dim iCount As Integer 'Counter used to determine if an odd number of parameter values is in the array (error)
    
        
 'Build string of parameter values
    iCount = 0
    sParamList = ""
    For i = LBound(vntVarList) To iBarIndex - 1
        If (i Mod 2) = 0 Then
            sParamList = sParamList & vntVarList(i)
            iCount = iCount + 1
        Else
            If VarType(vntVarList(i)) <> 8 Then
                sParamList = sParamList & " = '" & str$(vntVarList(i)) & "' | "
            Else
                sParamList = sParamList & " = '" & vntVarList(i) & "' | "
            End If
            iCount = iCount + 1
        End If
    Next
    
'Check for error
    If (iCount Mod 2) = 1 Then
        iIsError = True
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "(DebugBuildParameterList) Error: " & Error(Err)
    iIsError = True
    Exit Sub
End Sub
Private Sub DebugBuildVariableList(iBarIndex As Integer, aVarList() As Variant, sVars As String, iIsError As Integer)
'**********************************************************************************************************
'Sub Name: DebugBuildVariableList
'
'Description: Builds a bar separated listed of a procedures local variables passed in as a variant array.
'             i.e.  vntVarList = ("a",5,"b","abc","|","q",10) becomes
'                                "q='10'"
'             Notice the items before the "|" are ignored because those values are parameters for the procedure
'
'Parameters:
'iBarIndex - Location in the vntVarList where the bar ("|") character is located
'vntVarList - array that contains a calling procedures variables (and parameters) names and values
'           and variable names and values.  The parameters are specified first and a bar ("|") character
'           is used to separate the variables.  i.e. "p1", p1, "p2", p2, "|", "v1", v2
'sParamList - a string with the list of local variables i.e. "a='5' | b = 'abc'"
'iIsError - Indicates whether there was an error.  This typically occurs if there was an odd
'           number of values.
'**********************************************************************************************************
On Error GoTo ErrHandler


'Declarations ****************************************
    Dim i As Integer 'array index counter
    Dim iCount As Integer '

'Loop through array and put variables names and values into a string
    sVars = ""
    iCount = 0
    For i = iBarIndex + 1 To UBound(aVarList)
        If (i Mod 2) <> 0 Then
            sVars = sVars & aVarList(i)
            iCount = iCount + 1
        Else
            If VarType(aVarList(i)) <> 8 Then
                sVars = sVars & " = '" & str$(aVarList(i)) & "' | "
            Else
                sVars = sVars & " = '" & aVarList(i) & "' | "
            End If
            iCount = iCount + 1
        End If
        DoEvents
    Next
    If (iCount Mod 2) = 1 Then
        iIsError = True
    End If

    DoEvents
    Exit Sub
    
ErrHandler:
    MsgBox "(DebugBuildVariableList) Error: " & Error(Err)
    iIsError = True
    Exit Sub
End Sub
Private Sub TimeDifference(LastTime As Single, ThisTime As Single, DiffResult As String)
    On Error GoTo ErrHandler
Rem    Dim Minutes As Integer
Rem    Dim Seconds As Integer
    
Rem    Minutes = DateDiff("m", ThisTime, LastTime)
Rem    Seconds = DateDiff("s", ThisTime, LastTime)
Rem    Seconds = Seconds - (Mintues * 60)
Rem    If Minutes <> 0 Then
Rem        DiffResult = CStr(Minutes) & " min " & CStr(Seconds) & " Secs"
Rem    Else
Rem        DiffResult = CStr(Seconds) & " Secs"
Rem    End If

    Rem Dim Difference As Single
    Rem Difference = ThisTime - LastTime
    DiffResult = CStr(Format$(ThisTime - LastTime, "0.########"))
    Exit Sub
ErrHandler:
    MsgBox "(TimeDifference) Error: " & Error(Err)
    Exit Sub
End Sub
Private Sub TimePassedGet(sTimePassed As String)

    On Error GoTo ErrHandler
    
    sTimePassed = CStr(Format$(Timer - DebugTimer, "###0.0#######"))
    
    Exit Sub
ErrHandler:
    MsgBox "(TimePassedGet) Error: " & Error(Err)
    Exit Sub
End Sub
Public Sub DebugLog(sExe As String, sModule As String, sProc As String, sOper As String, tDR As cDebugLog.DebugRec, vntList() As Variant)
'**********************************************************************************************************
'Sub Name: DebugLog
'
'Description: Writes the given log information (sExe, sModule, etc..) to a log file and/or to a grid window.
'
'Parameters:
'sExe - Name of calling executable or DLL
'sModule - Name of calling VB module (FORM or BAS module)
'sProc - Name of calling procedure
'ssOper - Name of operation being performed within procedure (i.e Begin, End, ..)
'tDR - Record that tracks the processing time spent in a module
'VarList - unlimited length list of values that contains the calling procedures parameter names and values
'           and variable names and values.  The parameters are specified first and a bar ("|") character
'           is used to separate the variables.  i.e. "p1", p1, "p2", p2, "|", "v1", v2
'**********************************************************************************************************

On Error GoTo ErrHandler

Dim iBarIndex As Integer    'location of the "|" character in then vntVarList array
Dim bBarExists As Boolean   'True if "|" character in the vntVarList array otherwise false
Dim sParams As String       'List of all the procedures parameters extracted from the vntVarList array
Dim sVars As String         'List of the procedures variables extracted from the vntVarList Array
Dim sMemory As String       'How much memory is in the system at this time
Dim sPing As String         'If another computer was pinged this contains the time it took for a response
Dim fCurrTTime As Single    '?
Rem x Dim vntList() As Variant    'Array of parameters and variables passed into this procedure
Dim sTimeDiff As String     'Amount of time elapsed in the calling procedure since it last called this procedure
Dim sTheTime As String      '
Dim sTimePassed As String   '
Dim iIsError As Integer     '
Dim tTheDebugMode As DebugModes     '
Dim tTempDebugMode As DebugModes    '


'Get DEBUG MODE for the calling procedure ***********************************************************
    If giModuleDebugModeCount = 0 And giDefaultDebugMode = None Then
        Exit Sub
    ElseIf giModuleDebugModeCount = 0 Then
        tTheDebugMode = giDefaultDebugMode
    Else
        DebugGetTheMode sExe, sModule, sProc, tTempDebugMode
        If tTempDebugMode = None And giDefaultDebugMode = None Then Exit Sub
        If tTempDebugMode = None Then
            tTheDebugMode = giDefaultDebugMode
        End If
    End If

   fCurrTTime = Timer
   On Error Resume Next
   'DoEvents
   If tDR.TimerTime = 0 Then
        sTimeDiff = "NA"
   Else
        TimeDifference tDR.TimerTime, fCurrTTime, sTimeDiff
   End If
    TimePassedGet sTimePassed
    tDR.TimerTime = fCurrTTime
    'DoEvents
    Rem X vntList = aVarList
    DebugGetBarIndex iBarIndex, bBarExists, vntList()
    iIsError = False
    If iBarIndex = 0 Then
        sParams = ""
        sVars = ""
    Else
        DebugBuildParameterList iBarIndex, vntList(), sParams, iIsError
        If bBarExists Then
            DebugBuildVariableList iBarIndex, vntList(), sVars, iIsError
        Else
            sVars = ""
        End If
    End If
    sMemory = DebugMemory()
    sPing = "" 'giPING_RESULT
    debugRecordAdd sExe, sModule, sProc, sOper, sTimePassed, sTimeDiff, sMemory, sPing, sParams, sVars, tTheDebugMode, iIsError
    giPING_RESULT = ""
    'DoEvents
    Exit Sub
ErrHandler:
    MsgBox "(DebugLog) Error: " & Error(Err)
    Exit Sub
End Sub
Public Sub DebugLogPing(Exe As String, Module As String, Proc As String, Oper As String, DR As DebugRec, PingName As String, vntList() As Variant)
On Error GoTo ErrHandler
'ParamArray VarList() As Variant
'    DebugLog Exe$, Module$, Proc$, Oper$, DR, ParamArray VarList()', "VarList()
    
Dim iBarIndex As Integer
Dim bBarExists As Boolean
Dim sParams As String
Dim sVars As String
Dim sMemory As String
Dim sPing As String
Dim fCurrTTime As Single
Dim vntCurrTime As Variant
'Dim vntList() As Variant

Dim TimeDiff As String
Dim TheTime As String
Dim sTimePassed As String
Dim iIsError As Integer
Dim tTheDebugMode As DebugModes
Dim tTempDebugMode As DebugModes


    Dim i As Integer

   'MsgBox "piIPCount: " & CStr(piIPCount)
    giPING_RESULT = "'" & PingName & "' name not found."
    If piIPCount > 0 Then
        For i = LBound(paIPArr) To UBound(paIPArr)
            If paIPArr(i).Name = PingName Then
                giPING_RESULT = PingName & ": " & vbIcmpSendEcho(paIPArr(i).IP, "255", "10")
                Exit For
            End If
        Next
    End If

    If giModuleDebugModeCount = 0 And giDefaultDebugMode = None Then
        Exit Sub
    ElseIf giModuleDebugModeCount = 0 Then
        tTheDebugMode = giDefaultDebugMode
    Else
        DebugGetTheMode Exe, Module, Proc, tTempDebugMode
        If tTempDebugMode = None And giDefaultDebugMode = None Then Exit Sub
        If tTempDebugMode = None Then
            tTheDebugMode = giDefaultDebugMode
        End If
    End If

   fCurrTTime = Timer
   vntCurrTime = Now
   On Error Resume Next
   DoEvents
   If DR.TimerTime = 0 Then
        TimeDiff = "NA"
   Else
        TimeDifference DR.TimerTime, fCurrTTime, TimeDiff
   End If
    TimePassedGet sTimePassed
    Rem DR.LastTime = vntCurrTime
    DR.TimerTime = fCurrTTime
    DoEvents
    'vntList = VarList
    DebugGetBarIndex iBarIndex, bBarExists, vntList()
    iIsError = False
    If iBarIndex = 0 Then
        sParams = ""
        sVars = ""
    Else
        DebugBuildParameterList iBarIndex, vntList(), sParams, iIsError
        If bBarExists Then
            DebugBuildVariableList iBarIndex, vntList(), sVars, iIsError
        Else
            sVars = ""
        End If
    End If
    sMemory = DebugMemory()
    sPing = giPING_RESULT
    debugRecordAdd Exe, Module, Proc, Oper, sTimePassed, TimeDiff, sMemory, sPing, sParams, sVars, tTheDebugMode, iIsError
    giPING_RESULT = ""
    DoEvents
    Exit Sub
ErrHandler:
    MsgBox "(DebugLog) Error: " & Error(Err)
    Exit Sub

End Sub
Private Sub debugRecordAdd(Exe$, Module$, Proc$, Oper$, ElapsedTime$, sTimeDiff As String, sMemory As String, sPing As String, sParams As String, sVars As String, tTheDebugMode As DebugModes, iIsError As Integer)

On Error GoTo ErrHandler
  
    If giDisplayMode = Both Or giDisplayMode = ToScreen Then
        frmLog.fgLog.AddItem ""
        frmLog.fgLog.Row = frmLog.fgLog.Rows - 1
        frmLog.fgLog.Col = 0
        frmLog.fgLog = str(frmLog.fgLog.Rows - 1)
        If iIsError Then
            frmLog.fgLog.CellBackColor = 255
        End If
        frmLog.fgLog.Col = cEXE
        frmLog.fgLog = Exe$
        frmLog.fgLog.Col = cMOD
        frmLog.fgLog = Module$
        frmLog.fgLog.Col = cPROCEDURE
        frmLog.fgLog = Proc$
        If tTheDebugMode <> Terse Then
            frmLog.fgLog.Col = cOPERATION
            frmLog.fgLog = Oper$
        End If
        frmLog.fgLog.Col = cELAPSED
        frmLog.fgLog = ElapsedTime$
        frmLog.fgLog.Col = cDIFFERENCE
        frmLog.fgLog = sTimeDiff
        frmLog.fgLog.Col = cPING
        frmLog.fgLog = sPing
        frmLog.fgLog.Col = cMEMORY
        frmLog.fgLog = sMemory
        If tTheDebugMode <> Terse Then
            frmLog.fgLog.Col = cPARAMETERS
            frmLog.fgLog = sParams
        End If
        If tTheDebugMode <> Terse And giDefaultDebugMode <> Medium Then
            frmLog.fgLog.Col = cVARIABLES
            frmLog.fgLog = sVars
        End If
    End If
    If giDisplayMode = Both Or giDisplayMode = LogFile Then
        Dim Line$
'        If tTheDebugMode = Verbose Then
            Line$ = Exe$ & " | " & Module$ & " | " & Proc$ & " | " & Oper$ & " | " & ElapsedTime & " | " & sTimeDiff & " | " & sMemory & "|" & sPing & "|" & sParams & sVars
'        ElseIf tTheDebugMode = Medium Then
'            Line$ = Exe$ & " | " & Module$ & " | " & Proc$ & " | " & Oper$ & " | " & ElapsedTime & " | " & sTimeDiff & " | " & sParams
'        ElseIf tTheDebugMode = Terse Then
'            Line$ = Exe$ & " | " & Module$ & " | " & Proc$ & " | " & ElapsedTime & " | " & sTimeDiff
'        End If
        genUpdateLog gsDebugLogPath, Line$
    End If

    Exit Sub
ErrHandler:
    MsgBox "(debugRecordAdd) Error: " & Error(Err)
    Exit Sub
End Sub

Private Function DebugMemory() As String

On Error GoTo ErrHandler
    Dim MS As MEMORYSTATUS
    Dim sSR As String
    Dim sGR As String
    Dim sUSR As String

    
    winGetMemory MS, sSR, sGR, sUSR
    DebugMemory = MS.dwAvailPhys & "/" & MS.dwTotalPhys & " - " & CStr(Int((MS.dwAvailPhys / MS.dwTotalPhys) * 100)) & "%"
    Rem DebugMemory = MS.dwTotalPhys & ":" & MS.dwAvailPhys & "-" & MS.dwTotalVirtual & ":" & MS.dwAvailVirtual & "-" & MS.dwTotalPageFile & ":" & MS.dwAvailPageFile '& " - " & sSR & ":" & sUSR & ":" & sGR
    Exit Function

ErrHandler:
    MsgBox "(DebugMemory) Error: " & Error(Err)
    Exit Function
End Function

