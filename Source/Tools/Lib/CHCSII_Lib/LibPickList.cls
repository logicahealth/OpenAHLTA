VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           mclsLibPickList/LibPickList.cls
'''
'''Description:      Pick List box - allows to display items, select single one
'''                  or cancel
'''
'''Modules Required:
'''   frmLibPickList
'''   basLibApp
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME        As String = "mclsLibPickList."

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''


'''Properties
Public psPromptName        As String
Public psTitle             As String
Public pbAllowCancel       As Boolean
Public plMaxCommentsLen    As Long
Public plMinCommentsLen    As Long
Public psCaptionAccept     As String
Public psCaptionCancel     As String
Public psNCID_AppClient    As String
Public psSubTitle          As String
Public psListHeader        As String

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private gfrmPickList       As frmLibPickList
Private mobjList           As mclsLibList
Public pobjCWComm          As Object

Public Property Set pobjList(objList As mclsLibList)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' objList
'''   List of items requested to be used
'''
'''Description:
'''
'''
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
   Const sPROC_NAME     As String = "Set pobjList"


On Error GoTo ErrHnd

'''<DSN
20    Set mobjList = objList
100   If Not objList Is Nothing Then
   
      End If
'''>DSN

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property


Public Property Get pobjList() As mclsLibList
   Set pobjList = mobjList
End Property


Private Sub Class_Initialize()
   Set mobjList = New mclsLibList
   pbAllowCancel = True
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Set mobjList = Nothing
   Set pobjCWComm = Nothing
End Sub
Public Function P_bSelect(Optional ByVal sInfo As String, Optional ByVal bAlwaysOnTop As Boolean = False) As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sInfo
'''   Optional additional information requested to be appended to the prompt message
''' bAlwaysOnTop
'''   Indicates whether prompt window should alway be visible to the user
'''
'''
'''Description:   Display current list for user selection.
'''               Select default(first selected item in the list)
'''               Set Selected Item as current list item
'''
'''Note:
'''
'''Warning:       If Cancel is not allowed form will not uload until OK is pressed
'''               Keys of all items in specified list must represent integer value
'''
'''Updates Globals:
'''
'''Returns:       True  - If OK was pressed
'''               False - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bSelect"

   Dim bResult As Boolean
   Dim objModalMsg      As New mclsLibModalMsg
   
On Error GoTo ErrHnd

      '''Make sure form is available
10    If gfrmPickList Is Nothing Then
20       Set gfrmPickList = New frmLibPickList
40       Call Load(gfrmPickList)
      End If

      If bAlwaysOnTop Then
100      With gfrmPickList
120         Call gobjLibApp.LibApp_AlwaysOnTop(.hwnd, .Top, .Height, .Width)
         End With
      End If
      
200   With gfrmPickList
220      .psCaptionAccept = Trim$(psCaptionAccept)
230      .psCaptionCancel = Trim$(psCaptionCancel)
260      .plMaxCommentsLen = plMaxCommentsLen
270      .plMinCommentsLen = plMinCommentsLen
280      .psSubTitle = psSubTitle
290      .psListHeader = psListHeader

800      Call objModalMsg.P_bInformCore(gsNCID_CW_APPLICATION, psNCID_AppClient, pobjCWComm)
801      bResult = .P_bSelect(mobjList, psPromptName, psTitle, sInfo, pbAllowCancel)
      End With
      
999   P_bSelect = bResult
   
Exit Function
ErrHnd:
   P_bSelect = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl & vbCrLf & psPromptName)

End Function


