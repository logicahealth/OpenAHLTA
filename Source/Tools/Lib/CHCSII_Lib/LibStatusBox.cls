VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibStatusBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        mclsLibStatusBox/LibStatusBox.frm
'''
'''Description:   Generic Display Progress Form
'''
'''
'''Note:          Uses frmLibStatus
'''
'''
'''
'''
'''Modules Required:
''' frmLibStatus
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "mclsLibStatusBox."

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mbLoaded        As Boolean

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mfrmStatus            As frmLibStatus
Private WithEvents mcmdCancel As CommandButton
Attribute mcmdCancel.VB_VarHelpID = -1

Public Event pevLibStatusCancel()

Public Sub P_Display(ByVal bShow As Boolean, _
                     Optional ByVal bAlwaysOnTop As Boolean = True, Optional lHWndParent As Long _
                    )
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' bShow         Indicates whether to show or hide the form
''' bAlwaysOnTop  Indicates whether dialog should always be visible to the user
''' lHWndParent   Indicates whether to set parent window for the status box
'''
'''Description:   Depending on bShow flag show or hide status form
'''
'''Note:          AlwaysOnTop is ignored if bShow is false
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_Display"
   
On Error GoTo ErrHnd

   mfrmStatus.plHwndParent = lHWndParent
   Call mfrmStatus.P_Display(bShow, bAlwaysOnTop)
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub


Public Sub P_Move(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lLeft         Requested left coordinate
''' lTop          Requested top coordinate
''' lWidth        Requested width of the form
'''
'''Description:   Move and size the form according to the specified parameters
'''
'''Note:
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_Move"
   
On Error GoTo ErrHnd
   
   With mfrmStatus
      .plWidth = lWidth '''Adjst width of the controls
      Call .Move(lLeft, lTop, lWidth, .Height)
   End With

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub

Public Property Let pbAllowCancel(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Update cancel option according to specified value
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pbAllowCancel"
   
On Error GoTo ErrHnd

   mfrmStatus.pbAllowCancel = bVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Get pbCancel() As Boolean
   pbCancel = mfrmStatus.pbCancel
End Property
Public Property Let pbDisplayCount(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Indicates whether progress count will be displayed
'''               as part of the count information message
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pbDisplayCount"
   
On Error GoTo ErrHnd

   mfrmStatus.pbDisplayCount = bVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Let pbProgressBar(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Indicate whether progress bar will be visible
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "pbDisplayCount"
   
On Error GoTo ErrHnd

   mfrmStatus.pbProgressBar = bVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Get plMaxVal() As Long
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Set target value for the progress bar
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plMaxVal"
   
On Error GoTo ErrHnd

   plMaxVal = mfrmStatus.prgProgress.Max
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Let plMaxVal(ByVal lVal As Long)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Set target value for the progress bar
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plMaxVal"
   
On Error GoTo ErrHnd

   mfrmStatus.plMaxVal = lVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Property
Public Property Get plValue() As Long
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   get current progress count
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plValue"

On Error GoTo ErrHnd
   
   plValue = mfrmStatus.plValue
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Let plValue(ByVal lVal As Long)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Update progress counter
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plValue"

On Error GoTo ErrHnd
   
   mfrmStatus.plValue = lVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Property
Public Property Let plWidth(ByVal lVal As Long)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Set width of the form to specified value
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "plWidth"
   
On Error GoTo ErrHnd

   mfrmStatus.plWidth = lVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Let psCaption(ByVal sVal As String)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Update form caption
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psCaption"
   
On Error GoTo ErrHnd

   mfrmStatus.psCaption = sVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Let psMsgCount(ByVal sVal As String)
'''************************************************************************
'''Called By:        Any
'''
'''Arguments:
'''
'''Description:      Update Information in the count label
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psMsgCount"
   
On Error GoTo ErrHnd

   mfrmStatus.psMsgCount = sVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Let psMsgReport(ByVal sVal As String)
'''************************************************************************
'''Called By:        Any
'''
'''Arguments:
'''
'''Description:      Update Information in the message label
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psMsgReport"
   
On Error GoTo ErrHnd

   mfrmStatus.psMsgReport = sVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Let psMsgWarning(ByVal sVal As String)
'''************************************************************************
'''Called By:        Any
'''
'''Arguments:
'''
'''Description:      Update Warning Information in the message label
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psMsgWarning"
   
On Error GoTo ErrHnd

   mfrmStatus.psMsgWarning = sVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Private Sub Class_Initialize()
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Make sure all required objects and properties are
'''               available and initialized
'''
'''Note:
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Class_Initialize"
   
On Error GoTo ErrHnd
   
   Set mfrmStatus = New frmLibStatus
   Call Load(mfrmStatus)
   Set mcmdCancel = mfrmStatus.cmdCancel
   mbLoaded = True

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   Set mfrmStatus = Nothing

End Sub
Private Sub Class_Terminate()
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Make sure status form is unloaded
'''
'''Note:          Loaded in the initialize event
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Class_Terminate"
   
On Error GoTo ErrHnd
   
   If mbLoaded Then
      Call Unload(mfrmStatus)
   End If
   
   Set mfrmStatus = Nothing
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   Set mfrmStatus = Nothing
   
End Sub
Public Property Get pbAllowCancel() As Boolean
   pbAllowCancel = mfrmStatus.pbAllowCancel
End Property
Public Property Let psMsgPrompt(ByVal sVal As String)
'''************************************************************************
'''Called By:        Any
'''
'''Arguments:
'''
'''Description:      Update Information in the message label
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psMsgPrompt"
   
On Error GoTo ErrHnd

   mfrmStatus.psMsgPrompt = sVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Get psMsgCount() As String
   psMsgCount = mfrmStatus.psMsgCount
End Property
Public Property Get psMsgReport() As String
   psMsgReport = mfrmStatus.psMsgReport
End Property

Public Property Get psMsgWarning() As String
   psMsgWarning = mfrmStatus.psMsgWarning
End Property

Public Property Get psMsgPrompt() As String
   psMsgPrompt = mfrmStatus.psMsgPrompt
End Property
Public Property Get psCaption() As String
   psCaption = mfrmStatus.psCaption
End Property

Private Sub mcmdCancel_Click()
'''************************************************************************
'''Called By:        None
'''Arguments:
'''
'''Description:      Inform client that cancel was clicked
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "mcmdCancel_Click"

On Error GoTo ErrHnd
   
   RaiseEvent pevLibStatusCancel
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub


