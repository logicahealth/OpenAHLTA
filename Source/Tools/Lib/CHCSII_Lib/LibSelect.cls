VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gclsLibSelect"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibSelect/LibSelect.cls
'''
'''Description:      Generic select prompts:
'''                  Select File, Folder, Drive, Font, Color, Printer...
'''
'''Modules Required:
'''
'''
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "gclsLibSelect."


''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mdlg                     As CommonDialog
Attribute mdlg.VB_VarHelpID = -1
Private mfrmPath                 As frmLibPath

Public Function LibSelect_bColor _
      (ByRef lColor As Long, Optional ByVal lDefault As Long = vbWhite, _
       Optional ByVal bAllowCustom As Boolean, Optional ByRef sErrMsg As String _
      ) As Long
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lColor        Return Value
''' lDefault      Default Value
''' bAllowCustom  Indicates whether custom colors are allowed for selection
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:      Initialize common dialog to display color
'''                  Return user selection
'''
'''Note:             If prompt was canceled sErrMsg will contain "Was canceled"
'''
'''Updates Globals:
'''Returns:          If color was selected
'''                     True
'''                  else
'''                     False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibSelect_bColor"
   
On Error GoTo ErrHnd

   sErrMsg = ""
   With mdlg
      If bAllowCustom Then
         .Flags = cdlCCRGBInit Or cdlCCFullOpen
      Else
         .Flags = cdlCCRGBInit Or cdlCCPreventFullOpen
      End If
      .Color = lDefault '''Set default
      Call .ShowColor
      lColor = .Color   '''Return selection
   End With
   
   LibSelect_bColor = True
   
Exit Function
ErrHnd:
   LibSelect_bColor = False
   If Err.Number = cdlCancel Then
      sErrMsg = gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_WAS_CANCELED)
   Else
      sErrMsg = Err.Description
      Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   End If
   
End Function
Public Function LibSelect_bDrive _
      (ByRef sDrive As String, Optional ByVal bSetDefault As Boolean = True, _
       Optional ByVal sPrompt As String, Optional ByVal bAsNetworkPath As Boolean = True, _
       Optional ByVal bAlwaysOnTop As Boolean = True, Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sDrive       Used to return result and depending on bSetDefault used
'''               as default Drive
''' bSetDefault   Indicates whether folder specified in sDrive should be used as default
''' sPrompt       Optional Prompt Message
''' bAsNetworkPath Indicates whether Mapped drives should be returned as network shares
'''                  e.g.  if true  - H: mapped to \\Venus\Home willbe returned as \\Venus\Home
'''                        If false - if true - H: mapped to \\Venus\Home willbe returned as H:
''' bAlwaysOnTop  Indicates whether prompt window should alway be visible to the user
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Initialize Drive pick list
'''               Allow user select one drive and return it to the caller
'''
'''Note:          If prompt was canceled sErrMsg will contain "Was canceled"
'''
'''Updates Globals:
'''Returns:          If Drive was selected
'''                     True
'''                  else
'''                     False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibSelect_bDrive"
   
On Error GoTo ErrHnd

   sErrMsg = ""
   If mfrmPath Is Nothing Then
      Set mfrmPath = New frmLibPath
   End If
   
   LibSelect_bDrive = mfrmPath.P_bSelect _
                     (sDrive, bSetDefault, sPrompt, bAsNetworkPath, True, bAlwaysOnTop, sErrMsg)
   
Exit Function
ErrHnd:
   LibSelect_bDrive = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sDrive & vbCrLf & sPrompt)

End Function
Public Function LibSelect_bFile _
      (ByRef sFilePathName As String, Optional ByRef sFileName As String, _
       Optional ByVal bSetDefault As Boolean = True, _
       Optional ByVal sFilter As String, _
       Optional ByVal lFlags As Long, _
       Optional ByVal sTitle As String = "Select File", _
       Optional ByRef bOpenReadOnly As Boolean, _
       Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sFilePathName Used to return Full Path name of the file and depending on bSetDefault used
'''               as default
''' sFileName     used to return Name of the file without path.
''' bSetDefault   Indicates whether specified file should be used as default
''' sFilter       Filter criteria for the pick list
''' lFlags        dialog options if not specified defaults
'''               to: cdlOFNFileMustExist Or cdlOFNNoChangeDir Or cdlOFNLongNames Or cdlOFNCreatePrompt
''' sTitle        title of the dialog. If not specified defaults to 'Open'
''' bOpenReadOnly Used to return Value of the check box 'open as read only' and
'''               as initial value for it.
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:
'''
'''Note:             If prompt was canceled sErrMsg will contain "Was canceled"
'''                  By Default Check box 'open as read only' is hidden
'''
'''Updates Globals:
'''
'''Returns:          If file was selected
'''                     True
'''                  else
'''                     False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibSelect_bFile"
      
On Error GoTo ErrHnd
   
   With mdlg
      If bSetDefault Then
         .FileName = sFilePathName
      End If
      
      If sFilter = "" Then
         .Filter = "All Files (*.*)|*.*|"
      Else
         .Filter = sFilter
      End If
      .FilterIndex = 0
      
      '''If not specified
      If lFlags = 0 Then
         '''use default
         lFlags = cdlOFNFileMustExist Or cdlOFNNoChangeDir Or cdlOFNLongNames Or cdlOFNCreatePrompt Or cdlOFNHideReadOnly
      End If
      '''If specificaly indicated Read Only
      If bOpenReadOnly Then
         '''Make sure the box is visible
         If lFlags And cdlOFNHideReadOnly Then
            lFlags = lFlags Xor cdlOFNHideReadOnly
         End If
         '''Check  the box
         lFlags = lFlags Or cdlOFNReadOnly
      End If
      .Flags = lFlags
      
      '''If specific title is requisted
      If Not sTitle = "" Then
         .DialogTitle = sTitle
      '''else use defaut: 'Open'
      End If
      
      .ShowOpen
      '''Return result of selection
      bOpenReadOnly = .Flags And cdlOFNReadOnly
      sFilePathName = .FileName
      sFileName = .FileTitle
      
   End With

   LibSelect_bFile = True
   
Exit Function
ErrHnd:
   LibSelect_bFile = False
   If Err.Number = cdlCancel Then
      sErrMsg = gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_WAS_CANCELED)
   Else
      sErrMsg = Err.Description
      Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   End If

End Function
Public Function LibSelect_bFolder _
      (ByRef sFolder As String, Optional ByVal bSetDefault As Boolean = True, _
       Optional ByVal sPrompt As String, Optional ByVal bAsNetworkPath As Boolean = True, _
       Optional ByVal bAlwaysOnTop As Boolean = True, Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sFolder       Used to return result and depending on bSetDefault used
'''               as default folder
''' bSetDefault   Indicates whether folder specified in sFolder should be used as default
''' sPrompt       Optional Prompt Message
''' bAsNetworkPath Indicates whether Mapped drives should be returned as network shares
'''                  e.g.  if true  - H: mapped to \\Venus\Home willbe returned as \\Venus\Home
'''                        If false - if true - H: mapped to \\Venus\Home willbe returned as H:
''' bAlwaysOnTop  Indicates whether prompt window should alway be visible to the user
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Initialize Folders pick list
'''               Allow user select one folder and return it to the caller
'''
'''Note:          If prompt was canceled sErrMsg will contain "Was canceled"
'''
'''Updates Globals:
'''Returns:          If folder was selected
'''                     True
'''                  else
'''                     False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibSelect_bFolder"
   
On Error GoTo ErrHnd

   sErrMsg = ""
   If mfrmPath Is Nothing Then
      Set mfrmPath = New frmLibPath
   End If
   
   LibSelect_bFolder = mfrmPath.P_bSelect _
                     (sFolder, bSetDefault, sPrompt, bAsNetworkPath, False, bAlwaysOnTop, sErrMsg)
   
Exit Function
ErrHnd:
   LibSelect_bFolder = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sFolder & vbCrLf & sPrompt)

End Function
Public Function LibSelect_bFont _
      (ByRef objFont As StdFont, _
       Optional ByVal bSetDefault As Boolean = False, Optional ByRef lColor As Long, _
       Optional ByVal lFlags As Long = cdlCFBoth Or cdlCFForceFontExist Or cdlCFLimitSize Or cdlCFEffects, _
       Optional ByVal dSizeMin As Double, Optional ByVal dSizeMax As Double, _
       Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' objFont       Used as default based bSetDefault switch and return values
''' bSetDefault   Indicates whether properties of objFont should be used as default
''' lColor        Font Color - Used as default based bSetDefault switch and return value
''' dSizeMin      Minimum Font size
''' dSizeMax      Maximum Font Size
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:
'''
'''Note:          If prompt was canceled sErrMsg will contain "Was canceled"
'''               Size range will be evaluated only if cdlCFLimitSize flag is specified
'''               Size range is not specified will default to 6 to 14
'''
'''Updates Globals:
'''Returns:          If font was selected
'''                     True
'''                  else
'''                     False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibSelect_bFont"
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   '''If not specified
   If dSizeMax + dSizeMin = 0 Then
      '''Set to default
      dSizeMin = 6
      dSizeMax = 14
   '''If minimum specified and maximum is not
   ElseIf dSizeMax = 0 Then
      '''Set maximum to no limit
      dSizeMax = dSizeMin + 100
   '''Else
      '''uset as specified invalid values will cause an error
   End If
   
   With mdlg
      If bSetDefault Then
         '''Set default properties
         .FontName = objFont.Name
         .FontSize = objFont.Size
         .FontBold = objFont.Bold
         .FontItalic = objFont.Italic
         .FontStrikethru = objFont.Strikethrough
         .FontUnderline = objFont.Underline
         .Color = lColor
      End If
      
      .Flags = lFlags
      .Min = dSizeMin
      .Max = dSizeMax
      
      Call .ShowFont
      '''Return result of selection
      objFont.Name = .FontName
      objFont.Size = .FontSize
      objFont.Bold = .FontBold
      objFont.Italic = .FontItalic
      objFont.Strikethrough = .FontStrikethru
      objFont.Underline = .FontUnderline
      lColor = .Color
   End With
   
   LibSelect_bFont = True
   
Exit Function
ErrHnd:
   LibSelect_bFont = False
   If Err.Number = cdlCancel Then
      sErrMsg = gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_WAS_CANCELED)
   Else
      sErrMsg = Err.Description
      '''<DEV Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   End If

End Function
Public Function LibSelect_sPrinter _
      (ByRef sPrinter As String, Optional ByVal bSetDefault As Boolean = False, _
       Optional ByRef bCancel As Boolean, Optional ByRef sErrMsg As String _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''   Default Value
''' bCancel       Indicates whether user have canceled the prompt
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:
'''
'''Note:             If prompt was canceled sErrMsg will contain "Was canceled"
'''
'''Updates Globals:
'''Returns:          If succcessful
'''
'''                  else
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibSelect_sPrinter"
      
On Error GoTo ErrHnd

   Call gobjLibMsg.LibMsg_lOKCancel(lMSGOK_CURRENTLY_NOT_AVAILABLE, , , sPROC_NAME)
   LibSelect_sPrinter = ""
   
Exit Function
ErrHnd:
   LibSelect_sPrinter = ""
   If Err.Number = cdlCancel Then
      sErrMsg = gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_WAS_CANCELED)
   Else
      sErrMsg = Err.Description
      Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   End If
End Function
Private Sub Class_Initialize()
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:      Set reference to common dialog control
'''
'''Note:
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "Class_Initialize"
      
On Error GoTo ErrHnd
   
   Set mdlg = Lib_dlg
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   
End Sub
Private Sub Class_Terminate()
   Set mdlg = Nothing
   Set mfrmPath = Nothing
   
End Sub

