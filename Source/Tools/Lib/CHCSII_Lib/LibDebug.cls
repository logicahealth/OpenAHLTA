VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        mclsLibDebug/clsLibDebug.cls
'''
'''Description:   Debug log class. Utilizes DLog
'''               This class will log debug message in its terminate event with default message "Exit".
'''               This allows reduce amount of trace code.
'''
'''Note:          To work properly - object of this class need to be declared within the scope of the
'''                  entity, been traced. E.g. If you want to trace single procedure object of this
'''                  class must be declared as local variable.
'''               This modules assumes that debug trace option is controlled by DebugOn - conditional
'''                  compilation switch
'''               Name of the Module and procedure will be saved for Exit message at the time of the first call
'''                  to Mtd_bLogMessage - assumption made that first call must be made within the scope
'''                  of the entity been monitored.
'''
'''Modules Required:
'''   LibApp
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME      As String = "mclsLibDebug."

Private Const mlMESSAGE_MAX_LEN  As Long = 500

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private msExitMessage            As String
Private msExitModuleName         As String
Private msExitProcName           As String

Private mbErr                    As Boolean


Public Function Mtd_bLogMessage _
      (ByVal sModuleName As String, ByVal sProcName As String, _
       Optional ByVal sMsg As String, Optional ByVal sExitMessage As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Class_Terminate or Any
'''
'''Arguments
''' sModuleName   Name of the caller module
''' sProcName     Name of the caller procedure
''' sMsg          Optional Entry message. If empty defaults to 'Enter'
''' sExitMessage  Optional Exit message. If empty defaults to 'Exit'
'''
'''
'''Description:   log specified debug message
'''               If Exit message is specified
'''                  store it - will be used in terminate event
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       True  - if successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = msMODULE_NAME & "Mtd_bLogMessage"

On Error GoTo ErrHnd

   If sMsg = "" Then
      sMsg = gsENTER
   End If
   '''Store exit message. Note: will be evaluated
   psExitMessage = sExitMessage
   '''If initial call
   If msExitModuleName = "" Then
      '''Store initial module and procedure name for exit message
      msExitModuleName = sModuleName
      msExitProcName = sProcName
   End If
   
   '''Log message
   Call gobjLibApp.LibApp_DebugLog(sModuleName, sProcName, sMsg)
   
Exit Function
ErrHnd:
   Mtd_bLogMessage = False
   '''This function can be called multiple times.
   '''Do not annoy user with error messages - just inform about the reason of debug log failure ones
   If Not mbErr Then
      mbErr = True
      Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   End If
   
End Function


Public Property Get psExitMessage() As String
'''************************************************************************
'''Called By:     Any
'''Arguments
'''Description:   Return caurrent value of the Exit message
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   
   psExitMessage = msExitMessage

End Property


Public Property Let psExitMessage(ByVal sValue As String)
'''************************************************************************
'''Called By:     Any
'''Arguments
''' sValue - New text of the exit message
'''
'''Description:   Evaluate specified value of the exit message and store it if valid
'''
'''Note:
'''
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = msMODULE_NAME & "psExitMessage"

On Error GoTo ErrHnd

   sValue = Trim$(sValue)
   If Len(sValue) > 0 Then
      If Len(sValue) > mlMESSAGE_MAX_LEN Then
         sValue = Left$(sValue, mlMESSAGE_MAX_LEN)
      End If
      msExitMessage = sValue
   End If
   
Exit Property
ErrHnd:
   If Not mbErr Then
      mbErr = True
      Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   End If

End Property





Private Sub Class_Terminate()
'''************************************************************************
'''Called By:     None
'''Arguments
'''
'''Description:   Log exit message using values stored during initial call
'''
'''Note:
'''
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = msMODULE_NAME & "Class_Terminate"

On Error GoTo ErrHnd

   If Trim$(msExitMessage) = "" Then
      msExitMessage = "Exit"
   End If
   Call Mtd_bLogMessage(msExitModuleName, msExitProcName, msExitMessage)
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Sub


