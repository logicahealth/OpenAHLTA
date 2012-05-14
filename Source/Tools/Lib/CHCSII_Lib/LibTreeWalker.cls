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
Attribute VB_Name = "mclsLibTreeWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           LibTreeWalker.cls - mclsLibTreeWalker
'''
'''Description:      Allows recursive scan of the folder
'''
'''
'''Suggested Use:
'''
'''Note:
'''
'''Modules Required: gclsLibApp
'''
'''Author         :  Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constansts Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME      As String = "mclsLibTreeWalker."
''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Public pbPopErr                  As Boolean

Private mbWalking                As Boolean
Private msFileLikeFilter()       As String
Private mbFileFilter             As Boolean
Private mbStop                   As Boolean
Private mlFileCount              As Long
Private mbRecursive              As Boolean
Private mbProcessing             As Boolean

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mobjFS                   As New Scripting.FileSystemObject

'''Last Chance to Cancel
Public Event pLTWBeforeWalk(ByRef bCancel As Boolean)
Public Event pLTWNextFile(ByVal sFolder As String, ByVal sFile As String, ByVal sShortPath As String)
Public Event pLTWAfterWalk(ByVal lFileCount As Long)
Private Function M_bWalk(ByVal sFolder As String, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:           P_bWalk
'''
'''Arguments:
'''
'''Description:         Perform recursive scan of the specified folder
'''                     Based on file filter criteria pass file information to the client
'''                     If file filter is specified
'''                        send file information to client only if file matches specified
'''                        criteria
'''                     else
'''                        send file info for all files
'''Note:
'''
'''Updates Globals:
'''Returns:
'''
'''Author:              Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "M_bWalk"

   Dim objFile             As Scripting.File
   Dim objCurrentFolder    As Scripting.Folder
   Dim objChildFolder      As Scripting.Folder
   Dim bRaise              As Boolean
   Dim lIdx                As Long
   
On Error GoTo ErrHnd
   
   If Not mbStop Then
      sErrMsg = ""
      Set objCurrentFolder = mobjFS.GetFolder(sFolder)
            
      For Each objFile In objCurrentFolder.Files
         If Not mbStop Then
            bRaise = False
            If mbFileFilter Then
               For lIdx = 0 To UBound(msFileLikeFilter)
                  If UCase$(objFile.Name) Like UCase$(msFileLikeFilter(lIdx)) Then
                     bRaise = True
                     Exit For
                  End If
               Next lIdx
            Else
               bRaise = True
            End If
            If bRaise Then
               '''Send file Info to client
               RaiseEvent pLTWNextFile(objCurrentFolder.Path, objFile.Name, objFile.ShortPath)
               mlFileCount = mlFileCount + 1
               '''<DSN may also send short path and name
            End If
         End If
      Next objFile
      If mbRecursive Then
         For Each objChildFolder In objCurrentFolder.SubFolders
            If Not mbStop Then
               If Not M_bWalk(objChildFolder.Path, sErrMsg) Then
                  Exit For
               End If
            End If
         Next objChildFolder
      End If
   End If
    
   M_bWalk = sErrMsg = ""
   
Exit Function
ErrHnd:
   M_bWalk = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sFolder, False, pbPopErr)

End Function


Public Function P_bWalk _
(ByVal sStartFolder As String, Optional ByRef bCancel As Boolean, Optional sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:           Any
'''
'''Arguments:
''' sStartFolder        Starting point (root) of the scan
''' bCancel             Indicates whether scan was stoped by the client
'''                     (mbStop Property was set to true
''' sErrMsg             Used to return Error information if any
'''
'''Description:         If not currently running
'''                        Initiate Scan
'''                     else
'''                        Inform caller
'''Note:
'''
'''Updates Globals:
'''Returns:             If was able to complete the scan
'''                        True
'''                     Else
'''                        False
'''
'''Author:              Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME           As String = "P_bWalk"

   Dim bResult                As Boolean
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   If Not mbWalking Then
      mbWalking = True
      mbStop = False
      mlFileCount = 0
      RaiseEvent pLTWBeforeWalk(mbStop)
      If Not mbStop Then
         '''If successfuly processed
         If M_bWalk(sStartFolder, sErrMsg) Then
            bResult = True
         End If
      End If
      '''If no errors but canceled was canceled
      If bResult And mbStop Then
         bResult = False
         bCancel = True
         sErrMsg = gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_WAS_CANCELED)
      End If
      RaiseEvent pLTWAfterWalk(mlFileCount)
      mbWalking = False
   Else
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsALREADY, gsRUNNING)
   End If
   
   P_bWalk = bResult
   
Exit Function
ErrHnd:
   P_bWalk = False
   sErrMsg = Err.Description
   mbWalking = False
   Call gobjLibApp.LibApp_Error _
        (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sStartFolder, False, pbPopErr)

End Function
Public Property Get pbRecursive() As Boolean
   pbRecursive = mbRecursive
End Property

Public Property Let pbRecursive(ByVal bVal As Boolean)
   mbRecursive = bVal
End Property

Public Sub pbStop()
   mbStop = True
End Sub

Public Property Get psFileLikeFilter() As String
'''************************************************************************
'''Called By:           Any
'''
'''Arguments:
''' sVal                Pipe delimited list of filters
'''
'''Description:         Returns current files filter
'''Note:
'''
'''Updates Globals:
'''Returns:             If filter is specified and no errors
'''                        Pipe delimited list of search patterns
'''                     Else
'''                        Empty string
'''
'''Author:              Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME           As String = "Get psFileLikeFilter"

   Dim bResult                As Boolean
   
On Error GoTo ErrHnd

   If mbFileFilter Then
      psFileLikeFilter = Join(msFileLikeFilter, "|")
   Else
      psFileLikeFilter = ""
   End If
   
Exit Property
ErrHnd:
   psFileLikeFilter = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Property

Public Property Let psFileLikeFilter(ByVal sVal As String)
'''************************************************************************
'''Called By:           Any
'''
'''Arguments:
''' sVal                Pipe delimited list of filters
'''
'''Description:         Convert specified list to array
'''Note:
'''
'''Updates Globals:
'''Returns:
'''
'''Author:              Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME           As String = "Let psFileLikeFilter"

   Dim bResult                As Boolean
   
On Error GoTo ErrHnd

   If Len(Trim$(sVal)) > 0 Then
      msFileLikeFilter = Split(sVal, "|")
      mbFileFilter = True
   Else
      Erase msFileLikeFilter
      mbFileFilter = False
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sVal)
   
End Property
Private Sub Class_Initialize()
   pbPopErr = True
   mbRecursive = True
End Sub
Private Sub Class_Terminate()
   Set mobjFS = Nothing
End Sub


