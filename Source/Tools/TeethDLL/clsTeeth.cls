VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTeeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oSingleton As New srvSVGComm.SingletonFactory
Private oSVGComm As srvSVGComm.clsSVGComm

'37476
'Public mbIsDxView As Boolean
Private meBrowserViewMode As eViewMode

Public frmTeethX As Form
Public imgXTeethX As ImgX

Public Event DrawingFinished()
Public Event DblClickMultiTooth(iToothId As Integer)
Public Event SingleClickMultiTooth(iToothId As Integer)
Public Event SelectionCountChange(bSingleSelected As Boolean)
Public Event ToothChanged(iToothIDNew As Integer, iToothIDOld As Integer, RefreshGrid As Boolean)
Public Event priSecondClick(currentToothID As Integer)
Public Event AllFinished()

' SVG
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function getfocus% Lib "user" ()
Private mbIsLoadingDx As Boolean
Private mbIsLoadingTx As Boolean
Private mbIsLoadingDx603a As Boolean

Public gobjShared  As CWShared
  
Private mlTeethControlHwnd As Long
Private mlParentTeethControlHwnd As Long

Private Function ChangeRedInDxViewChildren(oParentElement As Object, ShowRed As Boolean)

'Not used anymore

Dim oTempElement As Object
    

    Set oTempElement = oParentElement
    
    'Get itself and it's siblings
    Do
        Debug.Print oSVGComm.TeethObject.TeethWinDx.printNode(oTempElement)
        
        'Get the children
        If oTempElement.hasChildNodes Then
            ChangeRedInDxViewChildren oTempElement.firstChild, ShowRed
        End If
        
        If oTempElement.hasAttributes Then
            If ShowRed Then
                If oTempElement.getAttribute("class") = "HideRed" Then
                    oTempElement.setAttribute "class", "FillRed"
                End If
            Else
                If oTempElement.getAttribute("class") = "FillRed" Then
                    oTempElement.setAttribute "class", "HideRed"
                End If
            End If
        End If
        
        Set oTempElement = oTempElement.nextSibling
        
    Loop Until oTempElement Is Nothing
    
End Function

Public Sub DisplaySplash()

    'Display the splash screen
    If Not frmTeethX Is Nothing Then
        frmSplash.DisplaySplash
    End If
    
End Sub

Public Sub HideSplash()
    
    'Hide the splash screen
    If Not frmTeethX Is Nothing Then
        frmSplash.HideSplash
    End If
    
End Sub


Public Sub RaiseAllFinished()

    'Finished drawing all the views
    '5-14-03
    RaiseEvent AllFinished
    
End Sub

Public Function ChangeRedInDxViewForSF603A(ShowRed As Boolean) As Boolean

'Change
'.FillRed{fill:red;stroke:red}
'to
'.FillRed{fill:none;stroke:none}
Dim oElement As Object
Dim oNewElement As Variant
Dim oEl As Variant
Dim sStyleTag As String
Dim a As Variant
Dim b As Variant
Dim c As Variant
Dim d As Variant

On Error GoTo ErrHnd

'sstyletag=osvgcomm.TeethObject.TeethWinDx.printNode(osvgcomm.TeethObject.TeethDocDx)
    
    'Loop through all elements and change any "class=FillRed" to "class=FillNone"
    
    'Get the xml string for the Style tag
'    Set oElement = oSVGComm.TeethObject.TeethDocDx.getElementById("style0")
'    Set oNewElement = oElement.sheet
'    Set a = oNewElement.cssRules
    
    If ShowRed Then
        Set oElement = oSVGComm.TeethObject.TeethDocDx.getElementById("svg0")
        oElement.removeChild oSVGComm.TeethObject.TeethDocDx.getElementById("mystyle")
    Else
        'Make all the reds disappear
        sStyleTag = "<style id='mystyle' type='text/css'><![CDATA[" _
            & " .FillRed{fill:none;stroke:none;}" _
            & " .OutlineRed{fill:none;stroke:none;stroke-width:4;}" _
            & " .OutlineRedMed{fill:none;stroke:none;stroke-width:2;}" _
            & " .OutlineRedSm{fill:none;stroke:red;stroke-width:1;}" _
            & " .FillRedDash{fill:none;stroke:none;stroke-width:4;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:2.73 4.55;}" _
            & " .RedOnly {fill:none;}" _
            & " ]]></style>"

'        'Make all the reds, blue
'        sStyleTag = "<style id='mystyle' type='text/css'><![CDATA[" _
'            & " .FillRed{fill:blue;stroke:blue;}" _
'            & " .OutlineRed{fill:none;stroke:blue;stroke-width:4;}" _
'            & " .OutlineRedMed{fill:none;stroke:blue;stroke-width:2;}" _
'            & " .OutlineRedSm{fill:none;stroke:blue;stroke-width:1;}" _
'            & " .FillRedDash{fill:none;stroke:blue;stroke-width:4;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:2.73 4.55;}" _
'            & " .RedOnly {fill:blue;}" _
'            & " ]]></style>"
        Set oElement = oSVGComm.TeethObject.TeethDocDx.getElementById("svg0")
        oElement.appendChild oSVGComm.TeethObject.TeethWinDx.parseXML(sStyleTag, oSVGComm.TeethObject.TeethDocDx)
    End If

'    Set oElement = oSVGComm.TeethObject.TeethDocDx
'
'    Debug.Print oSVGComm.TeethObject.TeethWinDx.printNode(oElement)
'    ChangeRedInDxViewChildren oElement, ShowRed
        
    
'Set b = a.Item(CLng(1))
'oNewElement.insertRule ".fillScott{fill:white;}", 2
'oNewElement.deleteRule (1)

'    sStyleTag = oSVGComm.TeethObject.TeethWinDx.printNode(oElement)
    
    ChangeRedInDxViewForSF603A = True
Exit Function

ErrHnd:
    MsgBox "Error in clsTeeth.ChangeRedInDxViewForSF603A: " & Err.Description

End Function

Public Function SetTeethControl(NewParentHwnd As Long)
        
    If frmTeethX Is Nothing Then
        Initialize
    End If
    
    If mlTeethControlHwnd = 0 Then
        'The control is on the frmteethx form
        mlTeethControlHwnd = frmTeethX.teethSVGX1.getHwnd
    End If
    
    SetParent mlTeethControlHwnd, NewParentHwnd
    
    mlParentTeethControlHwnd = NewParentHwnd
        
End Function

Public Function SetTeethControlForReports(NewParentHwnd As Long)
        
    'This sub is for dataviews
    
    If frmTeethX Is Nothing Then
        Initialize
    End If
    
    If mlTeethControlHwnd = 0 Then
        'The control is on the frmteethx form
        mlTeethControlHwnd = frmTeethX.teethSVGX1.getHwnd
    End If
    
    'Remove all red from the Dx View
'    oSVGComm.TeethObject.TeethDocDx
    
    frmTeethX.ShowTeethForReports
    SetParent mlTeethControlHwnd, NewParentHwnd
    
    mlParentTeethControlHwnd = NewParentHwnd
    
End Function

Public Sub Initialize()
'   oSVGComm.PatientID = 999 '87214
'   oSVGComm.DentalEncounterID = 999
'   oSVGComm.Age = 18
'   Debug.Print Now

   InitializeObjects
   If frmTeethX Is Nothing Then
      Set frmTeethX = New frmTeeth
      Load frmTeethX
      If gobjShared.CmdLineSwitches("DENTALSHOWTEETH") = "DENTALSHOWTEETH" Then frmTeethX.Show
      Set frmTeethX.oParentClsTeeth = Me
      'frmTeethX.teethSVGX1.Visible = False
   End If
   
   
   'frmTeeth.Show
   'frmTeeth.showTeeth
   'frmTeethX.teethSVGX1.Visible = False
   mbIsLoadingDx = True
   Set oSVGComm.TeethObject.TeethDocDx = Nothing
   Set oSVGComm.TeethObject.TeethWinDx = Nothing
   
   '37476
   BrowserViewMode = eDxView
   'setDxTxMode True
   
   loadTeethDx
   
   'frmTeethX.teethSVGX1.Visible = True
'   ShowTeeth
'   setDxTxMode True
End Sub
Public Function getChildHwnd() As Long
  If frmTeethX Is Nothing Then
     Initialize
   End If
   getChildHwnd = frmTeethX.teethSVGX1.getHwnd
End Function
Public Function getTeethPicture(iZoomFactor As Double) As IPictureDisp
   frmTeethX.Show
   frmTeethX.WindowState = vbNormal
   DoEvents
   imgXTeethX.Import.FromDeviceContext frmTeethX.teethSVGX1.getHdc, 50, 20, frmTeethX.teethSVGX1.ScaleWidth - 60, frmTeethX.teethSVGX1.ScaleHeight - 30
   imgXTeethX.Effects.Resize imgXTeethX.Width / iZoomFactor, imgXTeethX.Height / iZoomFactor
   Set getTeethPicture = imgXTeethX.Export.ToPicture
'   frmTeethX.Hide
    frmTeethX.WindowState = vbMinimized
End Function

Public Function getParentHwnd() As Long
   getParentHwnd = frmTeethX.elTeeth.hWnd
End Function
Public Sub ReleaseTeethControl()

    If frmTeethX Is Nothing Then
        '5-14-03 'Don't need to init if the control is nothing
        Exit Sub
        'Initialize
    End If
    
    If mlTeethControlHwnd = 0 Then
        'The control is on the frmteethx form
        mlTeethControlHwnd = frmTeethX.teethSVGX1.getHwnd
    End If
    
    'Do not remove
    '37476
    frmTeethX.teethSVGX1.showBrowser eViewMode.eDxView
    frmTeethX.teethSVGX1.showBrowser eViewMode.eTxView
    frmTeethX.teethSVGX1.showBrowser eViewMode.eDx603aView
    'frmTeethX.teethSVGX1.showBrowser True
    'frmTeethX.teethSVGX1.showBrowser2 True
    
    If mlParentTeethControlHwnd <> frmTeethX.elTeeth.hWnd Then
        'The form is not on frmTeethX so put it there
        SetParent mlTeethControlHwnd, frmTeethX.elTeeth.hWnd
        mlParentTeethControlHwnd = frmTeethX.elTeeth.hWnd
    End If

End Sub

'Public Sub setDxTxMode(bIsDx As Boolean)
'
'    mbIsDxView = bIsDx
'    frmTeethX.ShowTeeth
'
'End Sub

Public Sub ShowTeeth()
   frmTeethX.ShowTeethControl
End Sub
Public Sub InitializeObjects()
   Set oSVGComm = oSingleton.Singleton
   Set imgXTeethX = New ImgX
   imgXTeethX.Initialize
   
   '37476
   meBrowserViewMode = eDxView
   'mbIsDxView = True
   
End Sub

Public Property Let BrowserViewMode(vData As eViewMode)
    
    'This replaced setDxTXMode
    
    meBrowserViewMode = vData
    frmTeethX.ShowTeeth

End Property

Public Property Get BrowserViewMode() As eViewMode

    BrowserViewMode = meBrowserViewMode
    
End Property

Private Sub Class_Initialize()
'>> SCR #22741;   Developer: Jon Mack 04/04/2002 09:43 PM
'SCR Description: Added switch to help debug patient change and graphics not loading.

 Set gobjShared = New CWShared

   
End Sub

Private Sub Class_Terminate()
'   TerminateObjects
End Sub
Public Sub TerminateObjects()

'   Set oSVGComm.TeethObject.TeethDocDx = Nothing
'   Set oSVGComm.TeethObject.TeethDocTx = Nothing
   
    Set oSVGComm = Nothing
    Set oSingleton = Nothing
    Set imgXTeethX = Nothing
   
    mlTeethControlHwnd = 0
'
    'Do not uncomment this
    'Set gobjShared = Nothing
   
    If Not frmTeethX Is Nothing Then
       Unload frmTeethX
       Set frmTeethX = Nothing
    End If
    
    Set gobjShared = Nothing
    
End Sub

Public Function loadTeethDx()
    ' this is to simulate graphic charting
    ' call svg server
    Dim S As String
          
    S = oSVGComm.getTeethHTMLDX()
    Call frmTeethX.teethSVGX1.loadHTMLDX(S)
    
        
End Function
Public Function loadTeethTx()
    ' this is to simulate graphic charting
    ' call svg server
    Dim S As String
          
    S = oSVGComm.getTeethHTMLTx()
    Call frmTeethX.teethSVGX1.loadHTMLTx(S)
    
        
End Function
Public Function loadTeethDx603a()
    ' this is to simulate graphic charting
    ' call svg server
    Dim S As String
          
    S = oSVGComm.getTeethHTMLDx603a()
    Call frmTeethX.teethSVGX1.loadHTMLDx603a(S)
            
End Function

Friend Sub Finished()

    If mbIsLoadingDx = True And mbIsLoadingTx = False And mbIsLoadingDx603a = False Then
        'Finished loading Dx so load Tx
        mbIsLoadingDx = False
        mbIsLoadingTx = True
        Set oSVGComm.TeethObject.TeethDocTx = Nothing
        Set oSVGComm.TeethObject.TeethWinTx = Nothing
        Debug.Print "frmTeethX.TxTimer.Enabled = True"
        frmTeethX.TxTimer.Enabled = True

    ElseIf mbIsLoadingDx = False And mbIsLoadingTx = True And mbIsLoadingDx603a = False Then
        'Finished loading Tx so load Dx603a
        mbIsLoadingTx = False
        mbIsLoadingDx603a = True
        Set oSVGComm.TeethObject.TeethDocDx603a = Nothing
        Set oSVGComm.TeethObject.TeethWinDx603a = Nothing
        Debug.Print "frmTeethX.Dx603aTimer.Enabled = True"
        frmTeethX.Dx603aTimer.Enabled = True

    ElseIf mbIsLoadingDx = False And mbIsLoadingTx = False And mbIsLoadingDx603a = True Then
        'Finished loading Tx so load Dx603a
        mbIsLoadingDx603a = False
        Debug.Print "frmTeethX.tmrFinished.enabled = True"
        frmTeethX.tmrFinished.Enabled = True

    Else
        Debug.Assert False
    '37476 - Commented this out
        'Added this to account for problems when the user gets ahead of the timers.
    '    mbIsLoadingDx = False
    '    mbIsLoadingTx = True
    '    Set oSVGComm.TeethObject.TeethDocTx = Nothing
    '    Set oSVGComm.TeethObject.TeethWinTx = Nothing
    '    Debug.Print "frmTeethX.TxTimer.Enabled = True"
    '    frmTeethX.TxTimer.Enabled = True
        
    End If
   
    'Debug.Print Now
    RaiseEvent DrawingFinished
  
End Sub
Public Sub ChangeToothView(RHV As Integer)
    frmTeethX.teethSVGX1.DblClickMultiTooth (RHV)
End Sub
Friend Sub DblClickMultiTooth(RHV As Integer)
  RaiseEvent DblClickMultiTooth(RHV)
End Sub
Friend Sub SingleClickMultiTooth(RHV As Integer)
  RaiseEvent SingleClickMultiTooth(RHV)
End Sub
Friend Sub SelectionCountChange(bSingleSelected As Boolean)
    RaiseEvent SelectionCountChange(bSingleSelected)
End Sub
Friend Sub ToothChanged(iToothIDNew As Integer, iToothIDOld As Integer, RefreshGrid As Boolean)
     oSVGComm.ToothObject.ToothChanged (iToothIDNew)
    RaiseEvent ToothChanged(iToothIDNew, iToothIDOld, RefreshGrid)
   
End Sub
Friend Sub priSecondClick(currentToothID As Integer)
    RaiseEvent priSecondClick(currentToothID)
End Sub

