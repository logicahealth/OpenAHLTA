VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSumOLEModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarModule            As Object
Private mvarsObjectName       As String
Private mvarbSummaryViewRunning As Boolean
Private mvarbOnSummary        As Boolean
Private mvarlSumLeftPercent   As Long
Private mvarlSumTopPercent    As Long
Private mvarlSumWidthPercent  As Long
Private mvarlSumHeightPercent As Long
Private mvarsApplicationNCID  As String
Private mvarsApplicationName  As String
Private mvarsOkTocloseView    As String
Private mvarsOkTochangePatient As String
Private mvarMenuItems         As Collection
Private mvarbChanged          As Boolean

Public Property Get MenuItems() As Collection:          Set MenuItems = mvarMenuItems: End Property
Public Property Let MenuItems(ByVal RHV As Collection): Set mvarMenuItems = RHV:                End Property

Public Property Get bChanged() As Boolean:               bChanged = mvarbChanged:           End Property
Public Property Let bChanged(ByVal RHV As Boolean):      mvarbChanged = RHV:                End Property

Public Property Get bSummaryViewRunning() As Boolean:          bSummaryViewRunning = mvarbSummaryViewRunning: End Property
Public Property Let bSummaryViewRunning(ByVal RHV As Boolean): mvarbSummaryViewRunning = RHV:                End Property

Public Property Get sOkTochangePatient() As String:          sOkTochangePatient = mvarsOkTochangePatient: End Property
Public Property Let sOkTochangePatient(ByVal RHV As String): mvarsOkTochangePatient = RHV:                End Property

Public Property Get sOkTocloseView() As String:            sOkTocloseView = mvarsOkTocloseView:       End Property
Public Property Let sOkTocloseView(ByVal RHV As String):   mvarsOkTocloseView = RHV:                  End Property
                            
Public Property Get Module() As Object:                    Set Module = mvarModule:                  End Property
Public Property Set Module(RHV As Object):                 Set mvarModule = RHV:                     End Property
                            
Public Property Get sObjectName() As String:               sObjectName = mvarsObjectName:             End Property
Public Property Let sObjectName(ByVal RHV As String):      mvarsObjectName = RHV:                     End Property

Public Property Get bOnSummary() As Boolean:               bOnSummary = mvarbOnSummary:               End Property
Public Property Let bOnSummary(ByVal RHV As Boolean):      mvarbOnSummary = RHV:                      End Property

Public Property Get lSumLeftPercent() As Long:             lSumLeftPercent = mvarlSumLeftPercent:     End Property
Public Property Let lSumLeftPercent(ByVal RHV As Long):    mvarlSumLeftPercent = RHV:                 End Property

Public Property Get lSumTopPercent() As Long:              lSumTopPercent = mvarlSumTopPercent:       End Property
Public Property Let lSumTopPercent(ByVal RHV As Long):     mvarlSumTopPercent = RHV:                  End Property

Public Property Get lSumWidthPercent() As Long:            lSumWidthPercent = mvarlSumWidthPercent:   End Property
Public Property Let lSumWidthPercent(ByVal RHV As Long):   mvarlSumWidthPercent = RHV:                End Property

Public Property Get lSumHeightPercent() As Long:           lSumHeightPercent = mvarlSumHeightPercent: End Property
Public Property Let lSumHeightPercent(ByVal RHV As Long):  mvarlSumHeightPercent = RHV:               End Property

Public Property Get sApplicationNCID() As String:          sApplicationNCID = mvarsApplicationNCID:   End Property
Public Property Let sApplicationNCID(ByVal RHV As String): mvarsApplicationNCID = RHV:                End Property

Public Property Get sApplicationName() As String:          sApplicationName = mvarsApplicationName:   End Property
Public Property Let sApplicationName(ByVal RHV As String): mvarsApplicationName = RHV:                End Property


Private Sub Class_Initialize()
'-- set defaults
    mvarbSummaryViewRunning = False
    mvarbOnSummary = False
    mvarlSumLeftPercent = 0
    mvarlSumWidthPercent = 0
    mvarlSumHeightPercent = 0
    mvarsApplicationName = "Unknown"
    mvarsOkTocloseView = "YES"
    mvarsOkTochangePatient = "YES"
    Set mvarMenuItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set mvarModule = Nothing
    Set mvarMenuItems = Nothing
End Sub


Public Function Resize(ByVal vlParentScaleWidth As Long, _
                       ByVal vlParentScaleHeight As Long) As Boolean
    Dim lLeft   As Long
    Dim lTop    As Long
    Dim lWidth  As Long
    Dim lHeight As Long

On Error GoTo errhandler

    If (mvarlSumLeftPercent + mvarlSumTopPercent + mvarlSumWidthPercent + mvarlSumHeightPercent = 0) _
       Or mvarModule Is Nothing Then
       Exit Function
    End If
            
            
    lLeft = (mvarlSumLeftPercent / 100) * vlParentScaleWidth
    lTop = (mvarlSumTopPercent / 100) * vlParentScaleHeight
    lWidth = (mvarlSumWidthPercent / 100) * vlParentScaleWidth
    lHeight = (mvarlSumHeightPercent / 100) * vlParentScaleHeight
    
    '---inform the OLE server it needs to resize the specified summary view
    mvarModule.ResizeView cwiPATIENT_SUMMARY_VIEW, lLeft, lTop, lWidth, lHeight
    Resize = True
    mvarbSummaryViewRunning = True
    Exit Function
    
errhandler:
    MsgBox "Unable to send a summary 'Resize' message to the clinical module: " & mvarsApplicationName & ".  This module could be in an 'unknown state'.", vbExclamation, "Health History"

End Function

Public Function CloseModule() As Boolean
    On Error Resume Next
    Dim objoleserver3 As ICHCSII_CoreAPI3
    If mvarModule Is Nothing Then Exit Function
    
    If TypeOf mvarModule Is ICHCSII_CoreAPI3 Then
        Set objoleserver3 = mvarModule
        If objoleserver3.CloseView2(cwiPATIENT_SUMMARY_VIEW) Then  'Ask to Close
            mvarbSummaryViewRunning = False
            CloseModule = True
        End If
        Set objoleserver3 = Nothing
    Else
        mvarModule.CloseView cwiPATIENT_SUMMARY_VIEW
        mvarbSummaryViewRunning = False
        CloseModule = True
    End If
End Function
