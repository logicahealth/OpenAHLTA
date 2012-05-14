VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibModalMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           mclsLibModalMsg/ModalMsg.cls
'''
'''Description:      Object for informing core about modal messages
'''
'''Modules Required:
'''
'''Author         :  Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mbMsgSent          As Boolean
Private msNCID_AppTarget   As String
Private msNCID_AppSource   As String '<DSN


''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME        As String = "mclsLibModalMsg."
Private mobjCWComm         As Object
Public Function P_bInformCore _
         (Optional ByVal sNCID_AppTarget As String = gsNCID_CW_APPLICATION, _
          Optional ByVal sNCID_AppSource As String, Optional ByVal objCWComm As Object _
         ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''
'''
'''
'''Description:   If Core comunication object is avaialble
'''                  Inform core or optional specified App that modal form is up
'''
'''Note:          Reset in terminate event
'''               Caller reference to objCWComm is preserved
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
   Const sPROC_NAME     As String = "P_bInformCore"

   
On Error GoTo ErrHnd

      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gsENTER)
      #End If
      '''If new reference is specified
20    If Not objCWComm Is Nothing Then
40       Set mobjCWComm = objCWComm
60    ElseIf mobjCWComm Is Nothing Then
80       Set mobjCWComm = gobjCWComm
      End If
   
100   If Not mobjCWComm Is Nothing Then
         '''If specified
120      If Not sNCID_AppSource = "" Then
140         msNCID_AppSource = sNCID_AppSource
         '''else use value set in Let psNCID_AppSource
         End If
         '''If target is not specified
         If sNCID_AppTarget = "" Then
            '''use NCID of the Core
            msNCID_AppTarget = gsNCID_CW_APPLICATION
         Else
            '''Use specified
            msNCID_AppTarget = sNCID_AppTarget
         End If
         #If LIB_LOG Then
            Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("Calling cwiSHOWING_MODAL_FORM Target/Source:", msNCID_AppTarget, msNCID_AppSource))
         #End If
300      Call mobjCWComm.Message(cwiSHOWING_MODAL_FORM, "", msNCID_AppTarget, msNCID_AppSource)
320      mbMsgSent = True
      Else
         #If LIB_LOG Then
            Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("cwiSHOWING_MODAL_FORM was not called. mobjCWComm is nothing. Target/Source:", msNCID_AppTarget, msNCID_AppSource))
         #End If
      End If
      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gsEXIT)
      #End If
   
      P_bInformCore = True
   
Exit Function
ErrHnd:
   P_bInformCore = False
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl & sNCID_AppSource)

End Function

Public Property Get pobjCWComm() As Object
   Set pobjCWComm = mobjCWComm
End Property

Public Property Set pobjCWComm(objCWComm As Object)
   Set mobjCWComm = objCWComm
End Property

Public Property Get psNCID_AppSource() As String
   psNCID_AppSource = msNCID_AppSource
End Property

Public Property Let psNCID_AppSource(ByVal sVal As String)
   msNCID_AppSource = sVal
End Property

Private Sub Class_Terminate()
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   If Core comunication object is available
'''                  Inform core that modal form is down
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
   Const sPROC_NAME     As String = "Class_Terminate"

On Error GoTo ErrHnd
   
   #If LIB_LOG Then
      Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gsENTER)
   #End If
   If Not mobjCWComm Is Nothing And mbMsgSent Then
      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("Calling cwiDONE_SHOWING_MODAL_FORM Target/Source:", msNCID_AppTarget, msNCID_AppSource))
      #End If
      Call mobjCWComm.Message(cwiDONE_SHOWING_MODAL_FORM, "", msNCID_AppTarget, msNCID_AppSource)
   Else
      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gobjLibStr.LibStr_sBuildSentence("cwiDONE_SHOWING_MODAL_FORM was not called: ", mbMsgSent, mobjCWComm Is Nothing, msNCID_AppTarget, msNCID_AppSource))
      #End If
   End If
   #If LIB_LOG Then
      Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, gsEXIT)
   #End If
   
   Set mobjCWComm = Nothing
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub


