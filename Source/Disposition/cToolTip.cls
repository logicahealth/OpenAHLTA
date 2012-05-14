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
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Initialization of New ClassNames
Private Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Type tagINITCOMMONCONTROLSEX
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be initialized.
End Type

' ToolTip Styles
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40 ' comctl32.dll v5.8 require

Private Const CW_USEDEFAULT = &H80000000

Private Const WS_POPUP = &H80000000

Private Const WM_USER = &H400

' ToolTip Messages
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_ADDTOOL = (WM_USER + 4)
Private Const TTM_DELTOOL = (WM_USER + 5)
Private Const TTM_NEWTOOLRECT = (WM_USER + 6)
Private Const TTM_GETTOOLINFO = (WM_USER + 8)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private Const TTF_IDISHWND = &H1
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hWnd        As Long
    uId         As Long
    cRect       As RECT
    hinst       As Long
    lpszText    As String
End Type

Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
'        lfFaceName As String
        lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum TTStyle
    ttStyleStandard = 1
    ttStyleBalloon = 2
End Enum

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long


Private Const WM_SETFONT = &H30

Private Const GWL_STYLE = (-16)
Private hTT As Long
'local variable(s) to hold property value(s)
Private mvarTipWidth As Long 'local copy
Private mvarDelayTime As Long 'local copy
Private mvarVisibleTime As Long 'local copy
Private mvarBkColor As Long 'local copy
Private mvarTxtColor As Long 'local copy
Private mvarStyle As TTStyle 'local copy
Private mvarFont As StdFont 'local copy


'''''''''''''''''''''''''''
'Public Const CB_SHOWDROPDOWN = &H14F
'Public Const LB_SETHORIZONTALEXTENT = &H194
'
'
'Public Declare Function SendMessage Lib "user32" _
'                  Alias "SendMessageA" (ByVal hWnd As Long, _
'                                        ByVal wMsg As Long, _
'                                        ByVal wParam As Long, _
'                                        ByVal lParam As Long) As Long
'
'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Public Type COMBOBOXINFO
'   cbSize As Long
'   rcItem As RECT
'   rcButton As RECT
'   stateButton  As Long
'   hwndCombo  As Long
'   hwndEdit  As Long
'   hwndList As Long
'End Type
'
'Public Const CB_GETDROPPEDCONTROLRECT = &H15F
'Public Const CB_GETITEMHEIGHT = &H154
'Public Const CB_SETDROPPEDWIDTH = &H160
'
'Public Declare Function GetComboBoxInfo Lib "user32.dll" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long

'''    'Init ToolTips '{{BOOKMARK Sample Code For ToolTips}}
'''    Set moToolTip = New cToolTip
'''
'''    With moToolTip
'''        'Props affecting tooltip
'''        .Style = ttStyleStandard
'''        .VisibleTime = 9000
'''
'''        'Specific To Combo Box
'''        .AutoSizeComboBoxDropDown cboModifiers1
'''        Set .mobjComboBox = Me.cboModifiers1
'''
'''        'Set The Tip For Each Item
'''        For i = 0 To ListCount
'''            moToolTip.SetToolTipItem i, CStr(GetExtraModiferList()(xtra_longdesc, i))
'''        Next
'''
'''    End With

''''''''''''''''''''''''


Private WithEvents mobjComboBox As VB.ComboBox
Attribute mobjComboBox.VB_VarHelpID = -1
Friend Sub SetCombo(ByRef obj As VB.ComboBox)
    Set mobjComboBox = obj
End Sub

Public Property Set Font(ByVal vData As StdFont)
  Dim lf As LOGFONT, hFont As Long, lHeight As Long
  lHeight = -MulDiv(vData.Size, GetDeviceCaps(GetDC(hTT), 90&), 72)
  lf.lfCharSet = vData.Charset
  lf.lfItalic = Abs(vData.Italic)
  lf.lfStrikeOut = Abs(vData.Strikethrough)
  lf.lfUnderline = Abs(vData.Underline)
  lf.lfWeight = vData.Weight
  Dim tmpArr() As Byte
  tmpArr = StrConv(vData.Name & Chr$(0), vbFromUnicode)
  Dim i As Integer, lArr As Long
  lArr = UBound(tmpArr)
  For i = 0 To lArr
      lf.lfFaceName(i) = tmpArr(i)
  Next i
  hFont = CreateFontIndirect(lf)
  SendMessageLong hTT, WM_SETFONT, hFont, 1&
  Set mvarFont = vData
End Property

Public Property Get Font() As StdFont
    If mvarFont Is Nothing Then
        Set mvarFont = New StdFont
    End If
    Set Font = mvarFont
End Property

Public Property Let TxtColor(ByVal vData As Long)
  mvarTxtColor = vData
  SendMessageLong hTT, TTM_SETTIPTEXTCOLOR, vData, 0&
End Property

Public Property Get TxtColor() As Long
  TxtColor = mvarTxtColor
End Property

Public Property Let BkColor(ByVal vData As Long)
  mvarBkColor = vData
  SendMessageLong hTT, TTM_SETTIPBKCOLOR, vData, 0&
End Property

Public Property Get BkColor() As Long
  BkColor = mvarBkColor
End Property

Public Property Let VisibleTime(ByVal vData As Long)
   mvarVisibleTime = vData
   SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, vData
End Property

Public Property Get VisibleTime() As Long
  VisibleTime = mvarVisibleTime
End Property

Public Property Let DelayTime(ByVal vData As Long)
    mvarDelayTime = vData
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_INITIAL, vData
End Property

Public Property Get DelayTime() As Long
    DelayTime = mvarDelayTime
End Property

Public Property Let TipWidth(ByVal vData As Long)
   mvarTipWidth = vData
   SendMessageLong hTT, TTM_SETMAXTIPWIDTH, 0, vData
End Property

Public Property Get TipWidth() As Long
   TipWidth = mvarTipWidth
End Property

Public Property Let Style(ByVal vData As TTStyle)
   Dim lStyle As Long
   mvarStyle = vData
   If hTT Then
      lStyle = GetWindowLong(hTT, GWL_STYLE)
      If vData = ttStyleBalloon Then lStyle = lStyle Or TTS_BALLOON
      If vData = ttStyleStandard And (lStyle And ttStyleBalloon) Then lStyle = lStyle Xor TTS_BALLOON
      SetWindowLong hTT, GWL_STYLE, lStyle
   End If
End Property

Public Property Get Style() As TTStyle
   Style = mvarStyle
End Property

Private Sub InitComctl32(dwFlags As Long)
   Dim icc As tagINITCOMMONCONTROLSEX
   On Error GoTo Err_OldVersion
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   InitCommonControlsEx icc
   On Error GoTo 0
   Exit Sub
Err_OldVersion:
   InitCommonControls
End Sub

Public Sub SetToolTipObj(objHwnd As Long, sTipText As String, Optional bCenter As Boolean = False)
    Dim TI As TOOLINFO
    With TI
        .hWnd = objHwnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        If bCenter Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uId = objHwnd
        .lpszText = sTipText
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub SetToolTipItem(nItem As Long, sTipText As String, Optional bCenter As Boolean = False)
    Dim TI As TOOLINFO, rc As RECT
    Dim objHwnd As Long
    Dim ItemHeight As Long
   
    ItemHeight = SendMessage(mobjComboBox.hWnd, CB_GETITEMHEIGHT, nItem, ByVal 0&)
    objHwnd = GetComboListHandle(mobjComboBox)
    rc.Bottom = rc.Top + ItemHeight
    rc.Left = 0
    rc.Right = mobjComboBox.Width \ Screen.TwipsPerPixelX
    rc.Top = (nItem - mobjComboBox.TopIndex) * ItemHeight
    
    With TI
         .cRect = rc
         .hWnd = objHwnd
         .uFlags = TTF_SUBCLASS
         If bCenter Then
             .uFlags = .uFlags Or TTF_CENTERTIP
         End If
         .uId = nItem
         .lpszText = sTipText
         .cbSize = Len(TI)
     End With
     
     SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub DelToolTip(objHwnd As Long, Optional nItem As Long = -1)
   Dim TI As TOOLINFO
   TI.hWnd = objHwnd
   TI.cbSize = Len(TI)
   If nItem < 0 Then TI.uId = objHwnd Else TI.uId = nItem
   SendMessage hTT, TTM_DELTOOL, 0, TI
End Sub

Public Sub AjustItemRect(objHwnd As Long, nItem As Long, lft As Long, tp As Long, rght As Long, btm As Long)
    Dim TI As TOOLINFO, rc As RECT
    With TI
        .hWnd = objHwnd
        .uId = nItem
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_GETTOOLINFO, 0&, TI
    rc.Bottom = btm
    rc.Left = lft
    rc.Right = rght
    rc.Top = tp
    TI.cRect = rc
    SendMessage hTT, TTM_NEWTOOLRECT, 0&, TI
End Sub

Private Sub Class_Initialize()

   InitComctl32 ICC_BAR_CLASSES
   hTT = CreateWindowEx(8, "tooltips_class32", 0&, TTS_NOPREFIX Or TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)
   TipWidth = 300
   BkColor = &HDFFFFF 'Soft Yellow/Std Tool Tip
   TxtColor = vbBlack
   DelayTime = 500
   VisibleTime = 2000
   Font.Size = 12
   Font.Name = "MS Sans serif"
   Me.Style = ttStyleStandard
       
   
End Sub

Private Sub Class_Terminate()
  If hTT Then DestroyWindow (hTT)
  If Not mvarFont Is Nothing Then DeleteObject ObjPtr(mvarFont)
End Sub

Private Sub mobjComboBox_DropDown()
    Dim rc As RECT
    Dim i As Long
    Dim ItemHeight As Long
    Dim ItemSelected As Long

  ItemSelected = mobjComboBox.ListIndex

  If ItemSelected < 0 Then ItemSelected = 0

  If mobjComboBox.ListCount = 0 Then Exit Sub
  
  mobjComboBox.TopIndex = ItemSelected
  ItemSelected = mobjComboBox.TopIndex

  For i = 0 To mobjComboBox.ListCount - 1
    rc.Left = 0
    rc.Right = mobjComboBox.Width \ Screen.TwipsPerPixelX
    ItemHeight = SendMessage(mobjComboBox.hWnd, CB_GETITEMHEIGHT, i, ByVal 0&)
    rc.Top = (i - ItemSelected) * ItemHeight
    rc.Bottom = rc.Top + ItemHeight

    Me.AjustItemRect GetComboListHandle(mobjComboBox), i, rc.Left, rc.Top, rc.Right, rc.Bottom
  Next

End Sub

Private Sub mobjComboBox_Scroll()
    Dim rc As RECT
    Dim i As Long
    Dim ItemHeight As Long

    For i = mobjComboBox.TopIndex To mobjComboBox.ListCount - 1
      rc.Left = 0
      rc.Right = mobjComboBox.Width \ Screen.TwipsPerPixelX
      ItemHeight = SendMessage(mobjComboBox.hWnd, CB_GETITEMHEIGHT, i, ByVal 0&)
      rc.Top = (i - mobjComboBox.TopIndex) * ItemHeight
      rc.Bottom = rc.Top + ItemHeight

      Me.AjustItemRect GetComboListHandle(mobjComboBox), i, rc.Left, rc.Top, rc.Right, rc.Bottom
    Next
End Sub


Friend Sub AutoSizeComboBoxDropDown(cmb As ComboBox)
    'Parameters: cmb - mobjComboBox control/object to perform the Autosize on

    Dim CurrentEntryWidth As Integer
    Dim PixelLength As Long
    Dim x As Integer
    Dim oFormFont As StdFont
    Dim iScaleMode As Integer 'find the longest string in the list portion of the mobjComboBox
        
    'temporarily set the form font to the combo box font
    'First cache the font
    Set oFormFont = cmb.Parent.Font
    
    'now set the combo box font to the form font
    Set cmb.Parent.Font = cmb.Font
    
    'temporarily change the ScaleMode of the form to Pixel
    'first cache the ScaleMode
    iScaleMode = cmb.Parent.ScaleMode
    
    'now set the ScaleMode to Pixel
    cmb.Parent.ScaleMode = vbPixels
    
    'find out the length in pixels of the longest string in the combo box
    For x = 0 To cmb.ListCount - 1
        CurrentEntryWidth = cmb.Parent.TextWidth(cmb.List(x))
        If CurrentEntryWidth > PixelLength Then
            PixelLength = CurrentEntryWidth
        End If
    Next
    
    'then add 10 pixels for a good measure (actually, to account for mobjComboBox margins)
    SetComboBoxDropDownWidth cmb.hWnd, PixelLength + 10
    
    'reset the ScaleMode to its original value
    cmb.Parent.ScaleMode = iScaleMode
    
    'reset to the original form font
    Set cmb.Parent.Font = oFormFont

End Sub



Private Sub SetComboBoxDropDownWidth(hWnd As Long, WidthPx As Long)
    'Parameters: hWnd - handle to the mobjComboBox (gotten through mobjComboBox.hWnd)
    '            WidthPx - width of the list portion of the combo box in pixels

    SendMessage hWnd, CB_SETDROPPEDWIDTH, WidthPx, 0

End Sub
Friend Function GetComboListHandle(ctl As ComboBox) As Long
Dim CBI As COMBOBOXINFO
   
   CBI.cbSize = Len(CBI)
   Call GetComboBoxInfo(ctl.hWnd, CBI)
   GetComboListHandle = CBI.hwndList
End Function


