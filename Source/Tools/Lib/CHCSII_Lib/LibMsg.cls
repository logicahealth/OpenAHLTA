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
Attribute VB_Name = "gclsLibMsg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibMsg/LibMsg.cls
'''
'''Description:   Standard Messages Utility
'''
'''Modules Required:
'''   LibApp
'''   LibStr
'''   LibTerms
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "gclsLibMsg."

'''To add new OK/Cancel message add member to this enumerator and corresponding case in LibMsg_sOKCancelText
Public Enum pnumMsgOKType
   lMSGOK_0_CUSTOM
   lMSGOK_UNDEFINED
   lMSGOK_NONE
   lMSGOK_UNKNOWN
   lMSGOK_ERROR_OCCURRED
   lMSGOK_FILE_NOT_FOUND
   lMSGOK_NOT_FOUND
   lMSGOK_ACCESS_DENIED
   lMSGOK_INSUFFICIENT_ACCESS_LEVEL
   lMSGOK_CURRENTLY_NOT_AVAILABLE
   lMSGOK_NOT_AVAILABLE
   lMSGOK_RESERVED_BY_ANOTHER_USER
   lMSGOK_FAILED_TO_OPEN
   lMSGOK_FAILED_TO_CLOSE
   lMSGOK_FAILED_TO_CREATE_OBJECT
   lMSGOK_FAILED_TO_READ
   lMSGOK_FAILED_TO_ADD
   lMSGOK_FAILED_TO_DELETE
   lMSGOK_FAILED_TO_UPDATE
   lMSGOK_FAILED_TO_FIND
   lMSGOK_FAILED_TO_SELECT
   lMSGOK_INVALID_REFERENCE
   lMSGOK_INVALID_OPTION
   lMSGOK_INVALID_ENTRY
   lMSGOK_UNDEFINED_OPTION
   lMSGOK_CURRENTLY_NOT_SUPPORTED
   lMSGOK_ENTRY_REQUIRED
   lMSGOK_WAS_CANCELED
   lMSGOK_IN_PROGRESS
   lMSGOK_CURRENTLY_ACTIVE
End Enum
'''To add new Yes/No/Cancel message add member to this enumerator and corresponding case in LibMsg_sYesNoCancelText
Public Enum pnumMsgYesNoType
   lMSGYESNO_NONE
   lMSGYESNO_WOULD_YOU_LIKE_TO_CONTINUE   '''Would you like to continue?
   lMSGYESNO_OVERWRITE_EXISTING_FILE      '''Overwrite existing file?
   lMSGYESNO_CONTINUE_ANYWAY              '''Continue anyway?
   lMSGYESNO_DELETE_SELECTED_ITEMS        '''Delete all selected item(s)?
   lMSGYESNO_EXIT_PROGRAM                 '''Exit Program?
   lMSGYESNO_SAVE_CURRENT_CHANGES         '''Save current changes
   lMSGYESNO_WOULD_YOU_LIKE_TO            '''Would you like to
   lMSGYESNO_EXIT_DIALOGUE
   lMSGYESNO_END_PROCESS
   'lMSGYESNO_
   'lMSGYESNO_
   'lMSGYESNO_
End Enum

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private msNCID_AppClient                  As String '<DSN
Public psClientTitle                      As String

'''Indicates that message functions should return default as selected option instead of prompting the user
Public gbSuppressMessages                 As Boolean
'''Indicates whether OK button should be a default for input form
Public pbLibMsgInputAcceptOnEnter         As Boolean

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Public pobjCWComm                         As Object
Private mfrmInput                         As frmLibInput
Public Sub LibMsg_DisplayText _
      (ByVal sTitle As String, ByVal sItemName As String, ByVal sText As String, _
       Optional ByVal sInfo As String, Optional ByVal bAlwaysOnTop As Boolean = False, _
       Optional ByVal sCmdCaption As String = "&Close" _
      )
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sTitle        Text that will appear in the title bar of the form. Default to Client App.Title if empty
''' sItemName     Specified name of the input item.
'''               If specified will be used as: Please enter <Prompt Name> [or Cancel (if cancel allowed)]>
'''               If empty Prompt message will contain sInfo text only
''' sText         Text requested to be displayed
''' sInfo         Optional Additional information. Appended to the prompt message
''' bAlwaysOnTop  Indicates whether prompt window should alway be visible to the user
''' sCmdCaption
'''   Caption for Accept button, defaults to "Close"
'''
'''Description:   Display specified text with specified associated information
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
   Const sPROC_NAME     As String = "LibMsg_DisplayText"

   
On Error GoTo ErrHnd

100   Call LibMsg_sInputStr _
      (sTitle, sItemName, sText, 0, False, 0, sInfo, bAlwaysOnTop, True, , sCmdCaption, True)
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sTitle)

End Sub


Public Function LibMsg_lOKCancel _
      (ByVal lMsgOK As pnumMsgOKType, Optional ByVal bCancel As Boolean = False, _
       Optional ByVal bDefaultOK As Boolean = True, Optional ByVal sItemName As String, _
       Optional sInfo As String, Optional sTitle As String, _
       Optional sHelpFile As String, Optional lHelpContext As Long _
      ) As VbMsgBoxResult
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lMsgOK - Requested message. Use pnumMsgOKType defined in this module
''' bCancel - Option - Indicates whether to include Cancel button. Defaults to OK only
''' bDefaultOK - Optional - If true OK will be default button. Defaults to OK
''' sItemName - Optional Name of the Item. If specified will preface message as <Item Name> - <Message>
''' sInfo - Optional additional information. If specified will appear after the message
''' sTitle - Optional Title. Appears in the Title bar of the message box. Defaults to Client App.Title
''' sHelpFile - Optional path name of the help file
''' lHelpContext - optional Optional. Help context number. If context is provided, helpfile must also be provided
'''
'''Description:   If suppress messages is not requested
'''                  Display specified message and return VbMsgBoxResult value corresponding
'''                  to selection made by the user
'''
'''Note:          See also LibMsg_lYesNoCancel
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If OK is pressed
'''                  vbOK
'''               If Cancel is pressed
'''                  vbCancel
'''               if Error
'''                  Default
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibMsg_lOKCancel"

   Dim lResult          As Long
   Dim lButtons         As Long
   Dim sMsg             As String
   Dim objModalMsg      As New mclsLibModalMsg
   
On Error GoTo ErrHnd

   '''Return Default if messages suppressed
   lResult = IIf(bDefaultOK, vbOK, vbCancel)
   
   If Not gbSuppressMessages Then
      '''Get text of the specified message
      sMsg = LibMsg_sOKCancelText(lMsgOK)
      '''Setup selection options
      lButtons = IIf(bCancel, vbOKCancel, vbOKOnly)
      '''If cancel option is not available
      If Not bCancel Then
         bDefaultOK = True '''default to OK button
      End If
      lButtons = lButtons + vbInformation + IIf(bDefaultOK, vbDefaultButton1, vbDefaultButton2)
      
      '''If Item name is specified
      sItemName = Trim$(sItemName)
      If Not sItemName = "" Then
         sMsg = sItemName & " - " & sMsg
      End If
      '''If Additional information is specified
      sInfo = Trim$(sInfo)
      If Not sInfo = "" Then
         sMsg = sMsg & vbCrLf & vbCrLf & sInfo
      End If
      '''If title is not specified
      sTitle = Trim$(sTitle)
      If sTitle = "" Then
         sTitle = psClientTitle
      End If
      '''If help reference is not specified
      sHelpFile = Trim$(sHelpFile)
      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("Calling P_bInformCore. Arguments: ", gsNCID_CW_APPLICATION, msNCID_AppClient, pobjCWComm Is Nothing))
      #End If
      Call objModalMsg.P_bInformCore(gsNCID_CW_APPLICATION, msNCID_AppClient, pobjCWComm)
      If sHelpFile = "" Then
         lResult = MsgBox(sMsg, lButtons, sTitle)
      Else
         lResult = MsgBox(sMsg, lButtons, sTitle, sHelpFile, lHelpContext)
      End If
   End If
   
   '''Return Result
   LibMsg_lOKCancel = lResult
   
Exit Function
ErrHnd:
   LibMsg_lOKCancel = IIf(bDefaultOK, vbOK, vbCancel)
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function


Public Function LibMsg_lYesNoCancel _
      (ByVal lMsgOK As pnumMsgOKType, ByVal lMsgYesNo As pnumMsgYesNoType, _
       Optional ByVal bCancel As Boolean = False, Optional ByVal lDefault As Long = vbNo, _
       Optional ByVal sItemName As String, Optional sInfo As String, Optional sTitle As String, _
       Optional sHelpFile As String, Optional lHelpContext As Long _
      ) As VbMsgBoxResult
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lMsgOK - Requested information message. Use pnumMsgOKType defined in this module
''' lMsgYesNo - Requested question message. Use pnumMsgYesNoType defined in this module
''' bCancel - Option - Indicates whether to include Cancel button. Defaults to Yes/No only
''' lDefault - Optional Default button. Defaults to vbNO
''' sItemName - Optional Name of the Item. If specified will preface message as <Item Name> - <Message>
''' sInfo - Optional additional information. If specified will appear after the message
''' sTitle - Optional Title. Appears in the Title bar of the message box. Defaults to Client App.Title
''' sHelpFile - Optional path name of the help file
''' lHelpContext - optional Optional. Help context number. If context is provided, helpfile must also be provided
'''
'''Description:   If suppress messages is not requested
'''                  Display specified message and return VbMsgBoxResult value corresponding
'''                  to selection made by the user
'''
'''Note:          See also LibMsg_lOKCancel
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If Yes is pressed
'''                  vbYes
'''               If No is pressed
'''                  vbNo
'''               If Cancel is pressed
'''                  vbCancel
'''               if Error
'''                  Default
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibMsg_lYesNoCancel"

   Dim lResult          As Long
   Dim lButtons         As Long
   Dim sMsg             As String
   Dim objModalMsg      As New mclsLibModalMsg
   
On Error GoTo ErrHnd
   
   Select Case lDefault
      Case vbYes
         lButtons = vbDefaultButton1
      Case vbNo
         lButtons = vbDefaultButton2
      Case vbCancel
         '''If cancel option is available
         If bCancel Then
            lButtons = vbDefaultButton3
         Else
            '''Default to NO
            lButtons = vbDefaultButton2
         End If
      Case Else
         lDefault = vbNo
         lButtons = vbDefaultButton2
   End Select
   
   '''Return Default if messages suppressed
   lResult = lDefault
   
   If Not gbSuppressMessages Then
      '''Get text of the specified message
      sMsg = LibMsg_sOKCancelText(lMsgOK)
      '''Add Question
      sMsg = sMsg & vbCrLf & vbCrLf & LibMsg_sYesNoCancelText(lMsgYesNo)
      
      '''Setup selection options. Note: default is already set
      lButtons = lButtons + IIf(bCancel, vbYesNoCancel, vbYesNo) + vbQuestion
      
      '''If Item name is specified
      sItemName = Trim$(sItemName)
      If Not sItemName = "" Then
         sMsg = sItemName & " - " & sMsg
      End If
      '''If Additional information is specified
      sInfo = Trim$(sInfo)
      If Not sInfo = "" Then
         sMsg = sMsg & vbCrLf & vbCrLf & sInfo
      End If
      '''If title is not specified
      sTitle = Trim$(sTitle)
      If sTitle = "" Then
         sTitle = psClientTitle
      End If
      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("Calling P_bInformCore. Arguments: ", gsNCID_CW_APPLICATION, msNCID_AppClient, pobjCWComm Is Nothing))
      #End If
      Call objModalMsg.P_bInformCore(gsNCID_CW_APPLICATION, msNCID_AppClient, pobjCWComm)
      '''If help reference is not specified
      sHelpFile = Trim$(sHelpFile)
      If sHelpFile = "" Then
         lResult = MsgBox(sMsg, lButtons, sTitle)
      Else
         lResult = MsgBox(sMsg, lButtons, sTitle, sHelpFile, lHelpContext)
      End If
   End If
   
   '''Return Result
   LibMsg_lYesNoCancel = lResult
   
Exit Function
ErrHnd:
   LibMsg_lYesNoCancel = lDefault
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function

Public Function LibMsg_sInputStr _
      (ByVal sTitle As String, ByVal sItemName As String, ByVal sDefault As String, _
       ByRef bCancel As Boolean, _
       Optional ByVal bAllowCancel As Boolean = True, _
       Optional ByVal lMaxLen As Long = 255, _
       Optional ByVal sInfo As String, _
       Optional ByVal bAlwaysOnTop As Boolean = False, _
       Optional ByVal bMultiLine As Boolean = True, _
       Optional ByRef bError As Boolean, _
       Optional ByVal sAcceptCaption As String = "O&K", Optional ByVal bReadOnly As Boolean _
      ) As String
'''************************************************************************
'''Called By:     LibMsg_DisplayText
'''
'''Arguments:
''' sTitle        Text that will appear in the title bar of the form. Default to Client App.Title if empty
''' sItemName     Specified name of the input item.
'''               If specified will be used as: Please enter <Prompt Name> [or Cancel (if cancel allowed)]>
'''               If empty Prompt message will contain sInfo text only
''' sDefault      Initial Value
''' bCancel       Indicates whether user have canceled the prompt
''' bAllowCancel  Optional Indicates whether cancel option is allowed
''' lMaxLen       Optional Indicates maximum number of characters allowed for input - Defaults to 255
''' sInfo         Optional Additional information. Appended to the prompt message
''' bAlwaysOnTop  Indicates whether prompt window should alway be visible to the user
''' bMultiLine    Indicates whether multiple lines are allowed for input
''' bError        Optional - Indicates whether error has occurred
''' sAcceptCaption
'''   Caption for Accept button, defaults to "O&K"
''' bReadOnly
'''   Indicates whether Display value could be modified
'''
'''Description:   Prompt for string input
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  text entered by the user or cancel
'''               else
'''                  empty string and bError parameter will return True
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibMsg_sInputStr"

   Dim sResult          As String
   Dim objModalMsg      As New mclsLibModalMsg
   
On Error GoTo ErrHnd

   bError = False
   bCancel = False
   
   sTitle = Trim$(sTitle)
   If sTitle = "" Then
      sTitle = psClientTitle
   End If
   
   If mfrmInput Is Nothing Then
      Set mfrmInput = New frmLibInput
   End If
   Call Load(mfrmInput)
   
   If bAlwaysOnTop Then
      Call gobjLibApp.LibApp_AlwaysOnTop(mfrmInput.hwnd, mfrmInput.Top, mfrmInput.Height, mfrmInput.Width)
   End If
   #If CHCSII_LOG Then
      Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("Calling P_bInformCore. Arguments: ", gsNCID_CW_APPLICATION, msNCID_AppClient, pobjCWComm Is Nothing))
   #End If
   Call objModalMsg.P_bInformCore(gsNCID_CW_APPLICATION, msNCID_AppClient, pobjCWComm)
   
   mfrmInput.pbAcceptOnEnter = pbLibMsgInputAcceptOnEnter
   mfrmInput.pbMultiLine = bMultiLine
   mfrmInput.cmdInput(mfrmInput.plCMD_IDX_OK).Caption = sAcceptCaption
   
   If mfrmInput.Mtd_bPrompt(sDefault, sItemName, sTitle, lMaxLen, sInfo, bAllowCancel, , bError, bReadOnly) Then
      sResult = sDefault
   Else
      bCancel = True
   End If
   
   '''Return Result
   LibMsg_sInputStr = sResult
   
Exit Function
ErrHnd:
   bError = True
   LibMsg_sInputStr = ""
   bCancel = True
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sTitle)

End Function

Public Function LibMsg_sOKCancelText(ByVal lMsgType As pnumMsgOKType) As String
'''************************************************************************
'''Called By:     LibMsg_lOKCancel or any
'''
'''Arguments:
''' lMsgType - Requested message. Use pnumMsgOKType defined in this module
'''
'''Description:   Return text of the message corresponding to specified message type
'''
'''Note:          See Also LibMsg_sYesNoCancelText
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  message corresponding to specified message type
'''               else
'''                  Error message
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibMsg_sOKCancelText"

   Dim sResult          As String
   
On Error GoTo ErrHnd

   Select Case lMsgType
      Case lMSGOK_UNDEFINED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsUNDEFINED, gsERROR)
      Case lMSGOK_NONE, lMSGOK_0_CUSTOM
         sResult = ""
      Case lMSGOK_UNKNOWN
         sResult = gobjLibStr.LibStr_sBuildSentence(gsUNKNOWN, gsERROR)
      
      Case lMSGOK_ERROR_OCCURRED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsERROR, gsOCCURRED)
         
      Case lMSGOK_ACCESS_DENIED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsACCESS, gsDENIED)
      
      Case lMSGOK_ACCESS_DENIED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsINSUFFICIENT, gsACCESS, gsLEVEL, " - ", gsACCESS, gsDENIED)
         
      Case lMSGOK_NOT_FOUND
         sResult = gobjLibStr.LibStr_sBuildSentence(gsNOT, gsFOUND)
         
      Case lMSGOK_FILE_NOT_FOUND
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFILE, gsNOT, gsFOUND)
      
      Case lMSGOK_FAILED_TO_OPEN
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsOPEN)
         
      Case lMSGOK_FAILED_TO_CLOSE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsCLOSE)
         
      Case lMSGOK_FAILED_TO_READ
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsREAD)
         
      Case lMSGOK_FAILED_TO_ADD
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsADD)
         
      Case lMSGOK_FAILED_TO_DELETE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsDELETE)
         
      Case lMSGOK_FAILED_TO_UPDATE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsUPDATE)
         
      Case lMSGOK_FAILED_TO_FIND
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsFIND)
      
      Case lMSGOK_FAILED_TO_SELECT
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsSELECT)
         
      Case lMSGOK_CURRENTLY_NOT_AVAILABLE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsCURRENTLY, gsNOT, gsAVAILABLE)
         
      Case lMSGOK_NOT_AVAILABLE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsNOT, gsAVAILABLE)
         
      Case lMSGOK_RESERVED_BY_ANOTHER_USER
         sResult = gobjLibStr.LibStr_sBuildSentence(gsCURRENTLY, gsRESERVED, gsBY, gsANOTHER, gsUSER)
      
      Case lMSGOK_INVALID_REFERENCE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsREFERENCE)
         
      Case lMSGOK_FAILED_TO_CREATE_OBJECT
         sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsCREATE, gsOBJECT)
         
      Case lMSGOK_INVALID_OPTION
         sResult = gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsOPTION)
         
      Case lMSGOK_UNDEFINED_OPTION
         sResult = gobjLibStr.LibStr_sBuildSentence(gsUNDEFINED, gsOPTION)
         
      Case lMSGOK_CURRENTLY_NOT_SUPPORTED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsCURRENTLY, gsNOT, gsSUPPORTED)
      
      Case lMSGOK_ENTRY_REQUIRED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsENTRY, gsREQUIRED)
         
      Case lMSGOK_INVALID_ENTRY
         sResult = gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsENTRY)
      
      Case lMSGOK_WAS_CANCELED
         sResult = gobjLibStr.LibStr_sBuildSentence(gsWAS, gsCANCELED)
         
      Case lMSGOK_IN_PROGRESS
         sResult = gobjLibStr.LibStr_sBuildSentence(gsIN, gsPROGRESS)
         
      Case lMSGOK_CURRENTLY_ACTIVE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsCURRENTLY, gsACTIVE)
         
      Case Else
         sResult = gobjLibStr.LibStr_sBuildSentence(gsUNDEFINED, gsMESSAGE, gsTYPE, ": ", CStr(lMsgType))
         
   End Select
   
   '''Return Result
   LibMsg_sOKCancelText = sResult
   
Exit Function
ErrHnd:
   LibMsg_sOKCancelText = gsERROR
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsTYPE & lMsgType)
   
End Function


Public Function LibMsg_sYesNoCancelText(ByVal lMsgType As pnumMsgYesNoType) As String
'''************************************************************************
'''Called By:     LibMsg_lOKCancel or any
'''
'''Arguments:
''' lMsgType - Requested question. Use pnumMsgYesNoType defined in this module
'''
'''Description:   Return text of the message corresponding to specified message type
'''
'''Note:          See Also LibMsg_sOKCancelText
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  message corresponding to specified message type
'''               else
'''                  Error message
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibMsg_sYesNoCancelText"

   Dim sResult          As String
   Dim bQmark           As Boolean
   
On Error GoTo ErrHnd

   bQmark = True
   Select Case lMsgType
      Case lMSGYESNO_NONE
         sResult = ""
      Case lMSGYESNO_WOULD_YOU_LIKE_TO_CONTINUE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsWOULD, gsYOU, gsLIKE, gsTO, gsCONTINUE)
      
      Case lMSGYESNO_WOULD_YOU_LIKE_TO
         sResult = gobjLibStr.LibStr_sBuildSentence(gsWOULD, gsYOU, gsLIKE, gsTO) & " "
         bQmark = False
      Case lMSGYESNO_OVERWRITE_EXISTING_FILE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsOVERWRITE, gsEXISTING, gsFILE)
      
      Case lMSGYESNO_CONTINUE_ANYWAY
         sResult = gobjLibStr.LibStr_sBuildSentence(gsCONTINUE, gsANYWAY)
   
      Case lMSGYESNO_DELETE_SELECTED_ITEMS    '''Delete all sleeted items?
         sResult = gobjLibStr.LibStr_sBuildSentence(gsDELETE, gsSELECTED, Trim$(gsITEM), gs_S_)
         
      Case lMSGYESNO_EXIT_PROGRAM                 '''Exit Program?
         sResult = gobjLibStr.LibStr_sBuildSentence(gsEXIT, gsPROGRAM)
      
      Case lMSGYESNO_SAVE_CURRENT_CHANGES         '''Save current changes
         sResult = gobjLibStr.LibStr_sBuildSentence(gsSAVE, gsCURRENT, gsCHANGES)
   
      Case lMSGYESNO_EXIT_DIALOGUE
         sResult = gobjLibStr.LibStr_sBuildSentence(gsEXIT, gsDIALOGUE)
         
      Case lMSGYESNO_END_PROCESS
         sResult = gobjLibStr.LibStr_sBuildSentence(gsEND, gsPROCESS)
         
      Case Else
         sResult = gobjLibStr.LibStr_sBuildSentence(gsUNDEFINED, gsMESSAGE, gsTYPE, ": ", CStr(lMsgType))
         
   End Select
   
   If sResult <> "" Then
      If bQmark Then
         sResult = sResult & "?"
      End If
   End If
   '''Return Result
   LibMsg_sYesNoCancelText = sResult
   
Exit Function
ErrHnd:
   LibMsg_sYesNoCancelText = gsERROR
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsTYPE & lMsgType)
   
Exit Function
End Function

Private Sub Class_Initialize()
   pbLibMsgInputAcceptOnEnter = True
End Sub

Public Property Get psNCID_AppClient() As String
   psNCID_AppClient = msNCID_AppClient
End Property

Public Property Let psNCID_AppClient(ByVal sVal As String)
   msNCID_AppClient = sVal
End Property

Private Sub Class_Terminate()
On Error Resume Next
   Set pobjCWComm = Nothing
   If Not mfrmInput Is Nothing Then
      Call Unload(mfrmInput)
   End If
   Set mfrmInput = Nothing
   
End Sub


