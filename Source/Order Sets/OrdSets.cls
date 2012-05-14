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
Attribute VB_Name = "clsOrdSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''************************************************************************
'''Module:              clsOrdSets/OrdSets.cls
'''
'''Description:         Collection of the order items with associated manipulation and processing routines
'''                     and (optional) GUI
'''
'''Modules Required:
''' frmOrderSets
'''
'''Note:                This object allows to perform all basic functions: Load template items,
'''                     adding and queue items submiting order items to OE for processing
'''                     without instantiating associated form
'''
'''Author         :     Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME         As String = "clsOrdSets."
'''A&P TABS
Private Const mlANP_TAB_ORDER_SETS   As Long = 2
Private Const mlANP_TAB_RAD          As Long = 7
Private Const mlANP_TAB_LAB          As Long = 6
Private Const mlANP_TAB_MED          As Long = 8
'''Status box position
Private Const mlSTATUS_WIDTH        As Long = 5600
'''Grid columns
Private Enum mnumCOL
   mlCOL_SELECT
   mlCOL_NAME
   mlCOL_DETAILS
   mlCOL_MODIFY
   mlCOL_SORT
   mlCOL_KEY
   mlNO_OF_COL
End Enum
'''Standalone Grid Columns
Private Enum enumStandaloneGridCols
   sgcDISPLAY_NAME
   sgcKEY
   sgcFULL_NAME
   sgcNO_OF_COL
End Enum

'''Submit Event
Private Const msOPT_MODIFY       As String = "M"

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mbFormIsLoaded              As Boolean
Private mlModifyCount               As Long     '''<DEV may not be needed
Private msOrderStatus               As String

Private mbProcessingBatch           As Boolean
Private mbCancelBatch               As Boolean
Private mbBatchInitiatedLocally     As Boolean
Private mbY2TDisplay                As Boolean   'Indicates whether the Y2TDisplay switch is activated
Private msLoadedTemplateId          As String

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

'''Batch control objects
Private mcolOrdSets                 As Collection  '''of objects of clsOrderSetItem class
Private WithEvents mobjStatusBox    As mclsLibStatusBox
Attribute mobjStatusBox.VB_VarHelpID = -1
Private WithEvents mobjSubmitTmr    As mclsLibTimer
Attribute mobjSubmitTmr.VB_VarHelpID = -1
'''OE Interface
Private mobjOEGetSMan               As Object
Private mobjOESMan                  As Object

'''GUI interface
Private mfrmOrdSets                 As frmOrderSets
Private WithEvents mcmdDeSelectAll  As CommandButton
Attribute mcmdDeSelectAll.VB_VarHelpID = -1
Private WithEvents mcmdSelectAll    As CommandButton
Attribute mcmdSelectAll.VB_VarHelpID = -1
Private WithEvents mcmdRefreshData  As CommandButton
Attribute mcmdRefreshData.VB_VarHelpID = -1
Private WithEvents mcmdSubmit       As CommandButton
Attribute mcmdSubmit.VB_VarHelpID = -1
Private WithEvents ucNavigation     As ucTemplateNavigation
Attribute ucNavigation.VB_VarHelpID = -1

Private WithEvents mcmdSaveOrders As CommandButton
Attribute mcmdSaveOrders.VB_VarHelpID = -1
Private WithEvents mvsgOrderItems   As vsFlexGrid
Attribute mvsgOrderItems.VB_VarHelpID = -1
Private WithEvents mvsgStandalone   As vsFlexGrid
Attribute mvsgStandalone.VB_VarHelpID = -1

Private Function M_bAddOrdItemToCollection _
   (objItem As Object, ByVal sKey As String, ByVal bQueue As Boolean, _
    ByRef sErrMsg As String _
   ) As Boolean
'''************************************************************************
'''Called By:     P_bLoadTemplateItems
'''               P_bAddOrdItem
'''
'''Arguments:
''' objItem       Order Item(type Rx, Lab or Med) requested to be added to the queue
''' sKey          Key of the specified queue Item - used for Mtd_bRemoveOrder
''' bQueue        Indicates whether Item is part of the queue (else - template item)
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Add specified Item to collection of order items
'''
'''Note:          This procedure will not add item to the grid
'''
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bAddOrdItemToCollection"

   Dim lLine            As Long
   Dim bResult          As Boolean
   Dim objOrdItem       As clsOrderSetItem
   Dim sSortedKey       As String
   
On Error GoTo ErrHnd
                                                                                          lLine = 10
   Set objOrdItem = New clsOrderSetItem
                                                                                          lLine = 20
   If objOrdItem.P_bSetOrderItem(objItem, sKey, bQueue, sErrMsg) Then
      '''Preselect if queue item
      objOrdItem.pbSelected = bQueue
      
      'JT BEGIN SCR 35654 - Template object now provides items in user specified order
      '                     So if it is not queued, we know it is a template and we will use the order provided by template
      '                     Otherwise we keep the behavior as it is currently for pushed orders with green switch
      If gbGreenList = True And bQueue Then
            sSortedKey = M_iSortNameToKey(objOrdItem)
            
            If Len(sSortedKey) = 0 Then
            
              Call mcolOrdSets.Add(objOrdItem, sKey)
              
            Else
            
              Call mcolOrdSets.Add(objOrdItem, sKey, sSortedKey)
            
            End If
      Else
            'This is either a template item or we are not in green mode
            'If it's a template item, the template object provides items in the user specified order
            Call mcolOrdSets.Add(objOrdItem, sKey)
      
      End If
      'JT END SCR 35654
      
      bResult = True
   End If
                                                                                          lLine = 30
   Set objOrdItem = Nothing
   
   M_bAddOrdItemToCollection = bResult
   
Exit Function
ErrHnd:
   M_bAddOrdItemToCollection = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   
End Function
Private Function M_iSortNameToKey(pobjOrdItem As clsOrderSetItem) As String
Dim objOrdItem As clsOrderSetItem

    For Each objOrdItem In mcolOrdSets
        If objOrdItem.psSortName = pobjOrdItem.psSortName Then
            If StrComp(objOrdItem.psName, pobjOrdItem.psName) = 1 Then
                M_iSortNameToKey = objOrdItem.psKey
                Exit For
            End If
        End If
    Next
    
   

End Function
Private Function AddOrdItemToGrid _
      (ByVal sKey As String, ByVal bSort As Boolean, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     P_bAddOrdItem
'''
'''Arguments:
''' sKey          Key of the specified queue Item - used for Mtd_bRemoveOrder
''' sErrMsg       Error information if any - to be returned byref
''' bSort         Indicates whether to sort the grid
'''
'''Description:   Add specified Item to the grid
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "AddOrdItemToGrid"

   Dim lLine            As Long
   Dim objOrdSetItem    As clsOrderSetItem
   Dim lNewRow          As Long
   
On Error GoTo ErrHnd
                                                                                          lLine = 10
   Set objOrdSetItem = mcolOrdSets.Item(sKey)
                                                                                          lLine = 20
   With objOrdSetItem
      Call mvsgOrderItems.AddItem _
           (.pbSelected & vbTab & _
            .psName & vbTab & _
            .psDetails & vbTab & _
            .pbModify & vbTab & _
            .psSortName & vbTab & _
            .psKey _
           )
   End With
                                                                                          lLine = 30
   With mvsgOrderItems
                                                                                          lLine = 40
      lNewRow = .Rows - 1
                                                                                          lLine = 50
      Set .Cell _
         (flexcpPicture, lNewRow, mlCOL_SELECT) = mfrmOrdSets.ilsOrdSet.ListImages _
         (objOrdSetItem.psIcon).Picture
                                                                                          lLine = 60
      .Cell(flexcpFontBold, lNewRow, mlCOL_SELECT, lNewRow, mlCOL_MODIFY) = Not objOrdSetItem.pbQueueItem
                                                                                          lLine = 70
      .Cell(flexcpFontUnderline, lNewRow, mlCOL_SELECT, lNewRow, mlCOL_MODIFY) = objOrdSetItem.pbQueueItem
                                                                                          lLine = 90
      If bSort Then
                                                                                          lLine = 100
        Call .Select(1, mlCOL_SORT)
                                                                                          lLine = 110
        .Sort = flexSortStringNoCaseDescending
      End If
   End With
   
   Set objOrdSetItem = Nothing
   
   AddOrdItemToGrid = True
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   AddOrdItemToGrid = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function


Private Function M_bBatchChangeAnPTab(ByVal lItemType As enumItemType, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     M_bBatchSubmitItem
'''
'''Arguments:
''' lItemType     Type of the Order Item (Med,Rad,Lab)
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Activate tab corresponding to the type of requested item
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchChangeAnPTab"

   Dim lLine            As Long
   
On Error GoTo ErrHnd
                                                                                             lLine = 10
   sErrMsg = ""
                                                                                             lLine = 20
   Select Case lItemType
      Case itMed
                                                                                             lLine = 30
         Call gobjAnPParent.changetab(mlANP_TAB_MED)
      Case itLab
                                                                                             lLine = 50
         Call gobjAnPParent.changetab(mlANP_TAB_LAB)
      Case itRad
                                                                                             lLine = 60
         Call gobjAnPParent.changetab(mlANP_TAB_RAD)
      Case Else
                                                                                             lLine = 70
         sErrMsg = LibStr_sBuildSentence(gsINVALID, gsITEM, gsTYPE, ":", lItemType)
   End Select
                                                                                             lLine = 80
   M_bBatchChangeAnPTab = sErrMsg = ""
   
Exit Function
ErrHnd:
   sErrMsg = LibStr_sBuildSentence(gsFAILED, gsTO, gsCHANGE, gsTAB)
   M_bBatchChangeAnPTab = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Private Function M_bBatchGetNextItem _
      (ByVal bModify As Boolean, ByRef sKey As String, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     M_bBatchSubmitItem
'''
'''Arguments:
''' bModify       Indicates whether to search for Item that need to submited with option to modify
''' sKey          Key of the requested Item
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Find next selected Item
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If requested Item was found
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchGetNextItem"

   Dim lLine            As Long
   Dim objOrdSetItem    As clsOrderSetItem
   
On Error GoTo ErrHnd
                                                                                                lLine = 10
   sErrMsg = ""
                                                                                                lLine = 20
   For Each objOrdSetItem In mcolOrdSets
                                                                                                lLine = 30
      If objOrdSetItem.pbSelected Then
                                                                                                lLine = 40
         If bModify Then
                                                                                                lLine = 50
            If objOrdSetItem.pbModify Then
                                                                                                lLine = 60
               sKey = objOrdSetItem.psKey
            End If
         ElseIf Not objOrdSetItem.pbModify Then
                                                                                                lLine = 70
            sKey = objOrdSetItem.psKey
         End If
      End If
                                                                                                lLine = 80
      '''If requested item found
      If Not sKey = "" Then
                                                                                                lLine = 90
         Exit For
      End If
   Next objOrdSetItem
                                                                                                lLine = 100
   M_bBatchGetNextItem = Not sKey = ""
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   M_bBatchGetNextItem = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Public Function M_bBatchLoadNextItem _
      (ByRef sErrMsg As String, Optional ByVal sKey As String = "") As Boolean
'''************************************************************************
'''Called By:     P_bSubmitOrders
'''               M_bBatchSubmitItem
'''               P_OrderStatus
'''Arguments:
''' sErrMsg       Error information if any - to be returned byref
''' sKey          Optional key of the specific item requested to be processed
'''
'''Description:   if specific item is not requested
'''                  Find and load next selected item into timer object
'''               else
'''                  load specified item into timer object
'''
'''Note:          See also mobjSubmitTmr_pLibTimerEvent
'''
'''Warning:
'''
'''Returns:       True   - Item was loaded
'''               False  - Else
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchStart"

   Dim bResult          As Boolean
   Dim lLine            As Long
   Dim sEventID         As String
   Dim sSubmitType      As String
   
On Error GoTo ErrHnd

   sErrMsg = ""
                                                                                    lLine = 10
   '''If Specific Item is not requested
   If sKey = "" Then
                                                                                    lLine = 20
'''<DSN
'M_bBatchGetNextItem can be modified to look for Rx items fist then lab then rad
'to emulate the sort order of the grid
'''>DSN
      '''First process items that was not requested to be modified
      '''If cannot find non-modify item
      If Not M_bBatchGetNextItem(False, sKey, sErrMsg) Then
                                                                                    lLine = 40
         '''Attempt to find modify item
         Call M_bBatchGetNextItem(True, sKey, sErrMsg)
      End If
   End If
                                                                                    lLine = 60
   '''If not specified and none found
   If Not sKey = "" Then
                                                                                    lLine = 80
      If mcolOrdSets.Item(sKey).pbModify Then
                                                                                    lLine = 90
         sSubmitType = msOPT_MODIFY
      End If
                                                                                    lLine = 100
      '''Initiate event that will trigger processing of the next item
      sEventID = mobjSubmitTmr.P_sAddEventRegularInterval("s", 0, sSubmitType, sKey)
                                                                                    lLine = 110
      '''If failed to initiate event
      If sEventID = sKey Then
                                                                                    lLine = 120
         bResult = True
      Else
                                                                                    lLine = 130
         sErrMsg = LibStr_sBuildSentence(gsFAILED, gsTO, gsINITIATE, gsSUBMISSION, gsEVENT)
      End If
   '''Else - No more selected items or get next failed, which will be reflected in ErrMsg
   End If
                                                                                    lLine = 200
   M_bBatchLoadNextItem = bResult
   
Exit Function
ErrHnd:
   M_bBatchLoadNextItem = False
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Private Function M_bBatchUpdateStatusReport() As Boolean
'''************************************************************************
'''Called By:     M_bBatchUpdateStatus
'''               P_OrderStatus
'''               mobjStatusBox_pevLibStatusCancel
'''Arguments:
'''
'''Description:   If any additional error information
'''               Update status report of the batch
'''
'''Note:          msOrderStatus is set in P_OrderStatus
'''Warning:
'''
'''Returns:       True  - if successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchUpdateStatusReport"

   Dim lLine            As Long

On Error GoTo ErrHnd
                                                                                             lLine = 10
   '''If previous item failed or was canceled
   If Not msOrderStatus = "" Then
                                                                                             lLine = 40
      If Not mobjStatusBox Is Nothing Then
                                                                                             lLine = 50
         mobjStatusBox.psMsgReport = mobjStatusBox.psMsgCount & " - " & msOrderStatus _
                                   & vbCrLf & mobjStatusBox.psMsgReport
      End If
                                                                                             lLine = 60
      msOrderStatus = ""
   End If
                                                                                             lLine = 100
   M_bBatchUpdateStatusReport = True
   
Exit Function
ErrHnd:
   M_bBatchUpdateStatusReport = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

'SCRs 44213 and 44214 - Initializes the standalone order sets grid
Private Function InitStandaloneGrid() As Boolean
'SCR 52899 - Changed to use Version 6 of VSFlexGrid to avoid exception fault being caused by Version 7

On Error GoTo ErrHnd

Dim lColIdx As Long

   'Only Visible with Y2TDISPLAY switch
   If Not mbY2TDisplay Then
      InitStandaloneGrid = True
      Exit Function
   End If
   
   With mvsgStandalone
      
      'Turn off grid drawing for quicker initialization
      .Redraw = False
      
      Call .Clear
      .Rows = 1
      .FixedRows = 1
      .Cols = sgcNO_OF_COL
      
      'Setup the visible name column
      .ColWidth(sgcDISPLAY_NAME) = 2900
      .ColAlignment(sgcDISPLAY_NAME) = flexAlignLeftTop
      .TextMatrix(0, sgcDISPLAY_NAME) = "My Standalone Order Sets"
      
      'Setup a hidden column to track the template orderset key value
      .TextMatrix(0, sgcKEY) = "Key"
      .ColHidden(sgcKEY) = True
            
      'Setup a hidden collumn for the template's full name
      .TextMatrix(0, sgcFULL_NAME) = "Full Name"
      .ColHidden(sgcFULL_NAME) = True
      
      'Setup general grid properties
      .AllowSelection = False
      .HighLight = flexHighlightWithFocus
      .SelectionMode = flexSelectionByRow
      .GridLines = flexGridNone
      .ExplorerBar = flexExSort
                                                                                 
      'print single black border line around column header
      .Select 0, sgcDISPLAY_NAME
      Call .CellBorder(vbBlack, 0, 1, 1, 1, 1, 1)
            
      .Redraw = True
      
   End With
   
   InitStandaloneGrid = True
   
Exit Function
ErrHnd:
   InitStandaloneGrid = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, "InitStandaloneGrid", vbNullString)

End Function

Private Function InitGrid() As Boolean
'''************************************************************************
'''Called By:     InitGUI
'''
'''Arguments:
'''
'''Description:   Initialize grid
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
'SCR 52899 - Changed to use Version 6 of VSFlexGrid to avoid exception fault being caused by Version 7

   Const sPROC_NAME     As String = "InitGrid"

   Dim lLine            As Long
   Dim lCol             As Long
   
On Error GoTo ErrHnd
                                                                                 lLine = 10
   With mvsgOrderItems
      Call .Clear
                                                                                 lLine = 20
      .Rows = 1
                                                                                 lLine = 30
      .FixedRows = 1
                                                                                 lLine = 40
      .Cols = mlNO_OF_COL
                                                                                 lLine = 50
      .ColWidth(mlCOL_SELECT) = 750
                                                                                 lLine = 60
      .ColWidth(mlCOL_NAME) = 4210
                                                                                 lLine = 70
      .ColWidth(mlCOL_DETAILS) = 4450
                                                                                 lLine = 80
      .ColWidth(mlCOL_MODIFY) = 700
                                                                                 lLine = 90
      .ColAlignment(mlCOL_SELECT) = flexAlignCenterTop
                                                                                 lLine = 100
      .ColAlignment(mlCOL_NAME) = flexAlignLeftTop
                                                                                 lLine = 110
      .ColAlignment(mlCOL_DETAILS) = flexAlignLeftTop
                                                                                 lLine = 120
      .ColAlignment(mlCOL_MODIFY) = flexAlignCenterTop
      '''>DEV
                                                                                 lLine = 140
      .TextMatrix(0, mlCOL_SELECT) = "Select"
                                                                                 lLine = 150
      .TextMatrix(0, mlCOL_NAME) = "Name"
                                                                                 lLine = 160
      .TextMatrix(0, mlCOL_DETAILS) = "Details"
                                                                                 lLine = 170
      .TextMatrix(0, mlCOL_MODIFY) = "Modify"
                                                                                 lLine = 180
      .TextMatrix(0, mlCOL_SORT) = "Sort"
                                                                                 lLine = 190
      .TextMatrix(0, mlCOL_KEY) = "Key"
                                                                                 lLine = 200
      .ColHidden(mlCOL_SORT) = True
                                                                                 lLine = 210
      .ColHidden(mlCOL_KEY) = True
                                                                                 lLine = 220
      .ColDataType(mlCOL_SELECT) = flexDTBoolean
                                                                                 lLine = 230
      .ColDataType(mlCOL_MODIFY) = flexDTBoolean
                                                                                 lLine = 240
      .ColWidthMax = 8000
                                                                                 lLine = 250
      .AllowSelection = False
                                                                                 lLine = 260
      .HighLight = flexHighlightWithFocus
                                                                                 lLine = 270
      .SelectionMode = flexSelectionByRow
                                                                                 lLine = 280
      .GridLines = flexGridNone
                                                                                 lLine = 290
      .ExplorerBar = flexExSort
                                                                                 lLine = 300
      '''print single black border line around column headers
      For lCol = mlCOL_SELECT To mlCOL_MODIFY
                                                                                 lLine = 320 + lCol
         Call .Select(0, lCol)
                                                                                 lLine = 340 + lCol
         Call .CellBorder(vbBlack, 0, 1, 1, 1, 1, 1)
      Next lCol
   End With
   
   InitGrid = True
   
Exit Function
ErrHnd:
   InitGrid = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function


Private Function InitGUI() As Boolean
'''************************************************************************
'''Called By:     P_bOpen
'''
'''Arguments:
'''
'''Description:   Make sure form and required controls are available
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "InitGUI"

   Dim lLine            As Long
   
On Error GoTo ErrHnd
                                                                                             
   'SCRs 44213 and 44214 - Determine if the yellow template switch is enable which activates standalone order sets functionality
   If gobjShared Is Nothing Then Set gobjShared = New CWShared
   mbY2TDisplay = CBool(Len(gobjShared.CmdLineSwitches("YT2DISPLAY")) > 0)
                                                                                             lLine = 20
   Set mfrmOrdSets = New frmOrderSets
                                                                                             lLine = 30
   Call Load(mfrmOrdSets)
                                                                                             lLine = 40
   With mfrmOrdSets
                                                                                             lLine = 50
      .Visible = False
                                                                                             
      Set .GuiController = Me
                                                                                             lLine = 60
      Set mcmdDeSelectAll = .cmdDeSelectAll
                                                                                             lLine = 70
      Set mcmdSelectAll = .cmdSelectAll
                                                                                             lLine = 80
      Set mcmdRefreshData = .cmdRefreshData
                                                                                             lLine = 90
      Set mcmdSubmit = .cmdSubmit
                                                                                             lLine = 100
      Set mvsgOrderItems = .vsgOrderItems
      Set mvsgStandalone = .vsgStandalone
      Set mcmdSaveOrders = .cmdSaveOrderSet
      
      Set ucNavigation = .ucTemplateNavigation1
            
      'JT BEGIN SCR 39463 - ADDED NPOE support and now always show the Top Frame since we've added the NPOE functionality which is always visible but enable/disabled
      If gbGreenList = False Then
         'Hide the templates selection control and frame
         .ucTemplateNavigation1.Visible = False
         '.NPOESearch.MoveControls 60, 360 'in case it is decided that the control should be repositioned
      Else
         'Show the templates selection control and frame
         .ucTemplateNavigation1.Visible = True
      End If
      
      'Setup NPOE Controls
      .NPOESearch.MoveControls .ucTemplateNavigation1.Left + .ucTemplateNavigation1.Width + 500, 360
      .NPOESearch.ShowControls
      
      'JT SCR 42844 - NPOE selection is disabled for Theater
      If Not .NPOESearch.IsSignableProvider And Not gobjShared.IsAppMode(modeTheater) Then
         .NPOESearch.EnableControls
      End If
      'JT END SCR 39463
      
      'SCR 44213 and 44214 - Show standalone order order sets with Y2TDISPLAY switch
      If mbY2TDisplay Then
         .vsgStandalone.Visible = True
         .vsgOrderItems.Left = 3000
         .vsgOrderItems.Width = .Width - (2 * .vsgStandalone.Left) - .vsgStandalone.Width - 65
      Else
         .vsgStandalone.Visible = False
         .vsgOrderItems.Left = 240
         .vsgOrderItems.Width = .Width - (2 * .vsgOrderItems.Left)
      End If
      
   End With
   
   '''If was able setup grid columns and headers
   mbFormIsLoaded = InitStandaloneGrid() And InitGrid()
   InitGUI = mbFormIsLoaded
   
Exit Function
ErrHnd:
   InitGUI = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Private Function M_bRemoveTemplateItems() As Boolean
'''************************************************************************
'''Called By:     P_bOpen
'''
'''Arguments:
'''
'''Description:   Prepare main collection
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bRemoveTemplateItems"

   Dim objItem          As clsOrderSetItem
   Dim lLine            As Long
   
On Error GoTo ErrHnd
                                                                                             lLine = 20
   If mcolOrdSets Is Nothing Then
                                                                                             lLine = 30
      Set mcolOrdSets = New Collection
   Else
                                                                                             lLine = 40
      '''Make sure all template Items are removed
      For Each objItem In mcolOrdSets
                                                                                             lLine = 50
         If Not objItem.pbQueueItem Then
                                                                                             lLine = 60
            Call mcolOrdSets.Remove(objItem.psKey)
         End If
      Next objItem
   End If
   
   M_bRemoveTemplateItems = True
   
Exit Function
ErrHnd:
   M_bRemoveTemplateItems = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   
End Function

Public Function P_bLoadTemplateItems() As Boolean
'''************************************************************************
'''Called By:     P_bOpen
'''               mcmdRefreshData_Click
'''
'''Arguments:
'''
'''Description:   Read and load template Items
'''
'''Note:          This procedure will not add items to the grid
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bLoadTemplateItems"

   Dim bResult          As Boolean
   Dim lLine            As Long
   Dim objTemplateOps   As CHCSII_TemplateClient.TemplateOps
   Dim objTemplate      As CHCSII_TemplateClient.Template
   Dim lIdx             As Long
   Dim objOrdItem       As Object
   Dim objMsPnt         As New mclsLibMsPnt
   Dim bolOKAdd         As Boolean          'SCR 22465  4/1/02  Phil Crowder
   
   Static lKey_static   As Long
   
On Error GoTo ErrHnd
   
   Call objMsPnt.Mtd_SetMousePointer(Screen)
   
   If M_bRemoveTemplateItems Then
                                                                                             lLine = 10
      Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
                                                                                             lLine = 20
      Call objTemplateOps.Initialize _
         (cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser)
                                                                                             lLine = 30
      Set objTemplate = objTemplateOps.GetCurrentEncounterTemplate(OrderListTemplateSection)
      If objTemplate.TemplateName = "No Encounter Template Loaded" Then ' Is Nothing Then
         'SCR 48941 - Attempt to retreive the default template
         Set objTemplate = objTemplateOps.GetDefaultTemplate()
      End If
                                                                                             lLine = 40
      If Not objTemplate Is Nothing Then
                                                                                             lLine = 50
         With objTemplate.OrderList.Items
                                                                                             lLine = 60
            For lIdx = 0 To .Count - 1
                                                                                                         
              'SCR 22465  4/1/02  Phil Crowder
              Set objOrdItem = .Items(lIdx)

              Select Case TypeName(objOrdItem)
                Case "RxType"
                  bolOKAdd = gobjUser.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Write)
                Case "LabType"
                  bolOKAdd = gobjUser.HasPrivilegeEx(Priv_Orders_Laboratory_Orders, Priv_Write)
                Case "RadType"
                  bolOKAdd = gobjUser.HasPrivilegeEx(Priv_Orders_Radiology_Orders, Priv_Write)
                Case Else
                  bolOKAdd = True
              End Select
                                                                                             
              If bolOKAdd Then
                                                                                              lLine = 70
                lKey_static = lKey_static + 1
                                                                                              lLine = 80
                                                                                              lLine = 90
                Call M_bAddOrdItemToCollection(objOrdItem, "T" & lKey_static, False, "")
                                                                                              lLine = 100
              End If
              
              Set objOrdItem = Nothing

            Next
         End With
      End If
                                                                                             lLine = 110
      Set objTemplateOps = Nothing
                                                                                             lLine = 120
      Set objTemplate = Nothing
                                                                                             lLine = 130
      bResult = True
   End If
   
   P_bLoadTemplateItems = bResult
   
Exit Function
ErrHnd:
   P_bLoadTemplateItems = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function
Public Function P_bLoadTemplateItem(rsTemplateId As String) As Boolean
'''************************************************************************
'''Called By:     ucTemplateNavigate
'''
'''Arguments:     rsTemplateId
'''
'''Description:   Read and load template Item from navigation
'''
'''Note:          This procedure will not add items to the grid
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Jonathan Mack
'''************************************************************************
   Const sPROC_NAME     As String = "P_bLoadTemplateItem"

   Dim bResult          As Boolean
   Dim lLine            As Long
   Dim objTemplateOps   As CHCSII_TemplateClient.TemplateOps
   Dim objTemplate      As CHCSII_TemplateClient.Template
   Dim lIdx             As Long
   Dim objOrdItem       As Object
   Dim objMsPnt         As New mclsLibMsPnt
   Dim bolOKAdd         As Boolean          'SCR 22465  4/1/02  Phil Crowder
   
   Static lKey_static   As Long
   
On Error GoTo ErrHnd
   
   Call objMsPnt.Mtd_SetMousePointer(Screen)
   
   If M_bRemoveTemplateItems Then
                                                                                             lLine = 10
      Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
                                                                                             lLine = 20
      Call objTemplateOps.Initialize _
         (cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser)
                                                                                             lLine = 30
      Set objTemplate = objTemplateOps.GetTemplate(rsTemplateId)
                                                                                             lLine = 40
      If Not objTemplate Is Nothing Then
                                                                                             lLine = 50
         msLoadedTemplateId = objTemplate.TemplateID
         
         With objTemplate.OrderList.Items
                                                                                             lLine = 60
            For lIdx = 0 To .Count - 1
                                                                                                         
              'SCR 22465  4/1/02  Phil Crowder
              Set objOrdItem = .Items(lIdx)

              Select Case TypeName(objOrdItem)
                Case "RxType"
                  bolOKAdd = gobjUser.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Write)
                Case "LabType"
                  bolOKAdd = gobjUser.HasPrivilegeEx(Priv_Orders_Laboratory_Orders, Priv_Write)
                Case "RadType"
                  bolOKAdd = gobjUser.HasPrivilegeEx(Priv_Orders_Radiology_Orders, Priv_Write)
                Case Else
                  bolOKAdd = True
              End Select
                                                                                             
              If bolOKAdd Then
                                                                                              lLine = 70
                lKey_static = lKey_static + 1
                                                                                              lLine = 80
                                                                                              lLine = 90
                Call M_bAddOrdItemToCollection(objOrdItem, "T" & lKey_static, False, "")
                                                                                              lLine = 100
              End If
              
              Set objOrdItem = Nothing

            Next
         End With
         
      Else
         msLoadedTemplateId = vbNullString
      End If
                                                                                             lLine = 110
      Set objTemplateOps = Nothing
                                                                                             lLine = 120
      Set objTemplate = Nothing
                                                                                             lLine = 130
      bResult = True
   End If
   
   P_bLoadTemplateItem = bResult
   
Exit Function
ErrHnd:
   P_bLoadTemplateItem = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

'The template selection might have changed - reload if it is different from the currently loaded template
Private Function TemplateSelected(ByVal sTemplateId As String, ByVal bForceLoad As Boolean)

On Error GoTo Error_Handler

   If msLoadedTemplateId <> sTemplateId Or bForceLoad Then
      If P_bLoadTemplateItem(sTemplateId) = True Then
         Call PopulateOrdersGrid
      End If
   End If
   
Error_Handler:
   If CBool(Err) Then Call HandleErrors(Err, msMODULE_NAME, "TemplateSelected")
   
End Function

Friend Function PopulateOrdersGrid() As Boolean
'''************************************************************************
'''Called By:     P_bOpen
'''               P_bRemoveOrdItem
'''               P_bClearQueue
'''               mcmdRefreshData_Click
'''               M_bSelectAll
'''Arguments:
'''
'''Description:   Populate the grid with all items
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
'SCR 52899 - Changed to use Version 6 of VSFlexGrid to avoid exception fault being caused by Version 7

   Const sPROC_NAME     As String = "PopulateOrdersGrid"

   Dim lLine            As Long
   Dim bResult          As Boolean
   Dim objOrdSetItem    As clsOrderSetItem
   
On Error GoTo ErrHnd

   bResult = True
   
   With mvsgOrderItems
                                                                                          lLine = 10
      .Redraw = False
                                                                                          lLine = 30
      .Rows = 1
                                                                                          lLine = 20
      For Each objOrdSetItem In mcolOrdSets
         bResult = False
                                                                                          lLine = 40
         Call AddOrdItemToGrid(objOrdSetItem.psKey, False, "")
         bResult = True
      Next objOrdSetItem
                                                                                          lLine = 50

'      'SCR 20092  Phil Crowder 1/21/02
      If .Rows > 0 And .Cols >= mlCOL_SORT Then Call .Select(0, mlCOL_SORT)

      'SCR 34954 & 35654 - Don't Sort Items in Grid - Use order defined in the template
      .Sort = flexSortNone
                                                                                          lLine = 70
      .Redraw = True
   End With
                                                                                          lLine = 80
   PopulateOrdersGrid = bResult
   
Exit Function
ErrHnd:
   PopulateOrdersGrid = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Private Function M_bSelectAll(ByVal bSelect As Boolean) As Boolean
'''************************************************************************
'''Called By:     mcmdSelectAll_Click
'''               mcmdDeSelectAll_Click
'''
'''Arguments:
''' bSelect       Indicates whether to select or deselect all items
'''
'''Description:   Select or diselect all items
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bSelectAll"
   
   Dim bResult          As Boolean
   Dim lLine            As Long
   Dim objOrdSetItem    As clsOrderSetItem
   
On Error GoTo ErrHnd
                                                                                          lLine = 10
   For Each objOrdSetItem In mcolOrdSets
                                                                                          lLine = 20
      objOrdSetItem.pbSelected = bSelect
      '''If deselect
      If Not bSelect Then
                                                                                          lLine = 30
         '''make sure that modified is unchecked
         objOrdSetItem.pbModify = False
      End If
   Next objOrdSetItem
                                                                                          lLine = 40
   If mbFormIsLoaded Then
                                                                                          lLine = 50
      bResult = PopulateOrdersGrid
   Else
      bResult = True
   End If
   
   M_bSelectAll = bResult
   
Exit Function
ErrHnd:
   M_bSelectAll = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Private Function M_bBatchSubmitItem _
       (ByVal sKey As String, ByVal bAllowModify As Boolean, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     mobjSubmitTmr_pLibTimerEvent
'''
'''Arguments:
''' sKey          Key of the specified queue Item - used for Mtd_bRemoveOrder
''' sErrMsg       Error information if any - to be returned byref
''' bAllowModify  Indicates whether Item need to submited with option to modify
'''
'''Description:   Submit specified Item to OE for processing
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchSubmitItem"

   Dim lLine                  As Long
   Dim objOrdSetItem          As clsOrderSetItem
   Dim sPrompt                As String
   Dim bItemSent              As Boolean
   Dim bForceModify           As Boolean
   Dim objMsPnt               As New mclsLibMsPnt
   Dim bResetOrderingProvider As Boolean
   
On Error GoTo ErrHnd

   sErrMsg = ""
                                                                                    lLine = 10
   Call objMsPnt.Mtd_SetMousePointer(Screen)
                                                                                    lLine = 20
   '''If item has to be modified
   If bAllowModify Then
                                                                                    lLine = 30
      '''Enable AnP GUI
      Call gobjAnPParent.UnlockAnP(msMODULE_NAME)
   Else
                                                                                    lLine = 40
      '''Disable AnP GUI
      Call gobjAnPParent.LockAnP(msMODULE_NAME)
   End If
                                                                                    lLine = 50
   '''If user did not cancel the batch
   If Not mobjStatusBox.pbCancel Then
                                                                                    lLine = 60
      Set objOrdSetItem = mcolOrdSets(sKey)
                                                                                    lLine = 70
      '''Make sure item is submitted only ones per user selection
      objOrdSetItem.pbSelected = False
                                                                                    lLine = 80
      objOrdSetItem.pbModify = False
                                                                                    lLine = 90
      '''Update Status information in the status form
      Call M_bBatchUpdateStatus(objOrdSetItem.psName, bAllowModify, objOrdSetItem.psIcon)
                                                                                    lLine = 100
      If M_bBatchChangeAnPTab(objOrdSetItem.plType, sErrMsg) Then
         
         'JT BEGIN SCR 39463 - Reset the ordering provider if this order is from a template and not being modified and for a non signing provider
         If Not mfrmOrdSets.NPOESearch.IsSignableProvider And mfrmOrdSets.NPOESearch.SelectedProviderIEN <> "" And _
            Not objOrdSetItem.pbModify And Not objOrdSetItem.pbQueueItem Then
            'Reset the provider IEN of the session manager is this is a non signable provider and notify submit to make call to update the order before sending
            mobjOESMan.OrderingProviderIen = mfrmOrdSets.NPOESearch.SelectedProviderIEN
            bResetOrderingProvider = True
         End If
                                                                                       lLine = 130
         '''Process an item
         bItemSent = objOrdSetItem.P_bSubmit(mobjOESMan, bAllowModify, bForceModify, mbCancelBatch, sErrMsg, bResetOrderingProvider)
         'JT END SCR 39463
                                                                                    lLine = 132
         msOrderStatus = sErrMsg
                                                                                    lLine = 140
         If bForceModify Then
                                                                                    lLine = 150
            '''Just in the case if event hangs -
            '''Make sure AnP is enabled to allow user click clear
            Call gobjAnPParent.UnlockAnP(msMODULE_NAME)
         End If
      Else
                                                                                    lLine = 190
         mbCancelBatch = True
      End If
   Else
                                                                                    lLine = 200
      mbCancelBatch = True
   End If
                                                                                    lLine = 300
   '''If still processing
   '''Note: Batch could be terminated during P_bSubmit above -
   '''      if OE message is returned within single thread (create new or error)
   If mbProcessingBatch Then
                                                                                    lLine = 320
      '''If failed to submit item and batch was not canceled
      If (Not bItemSent) And (Not mbCancelBatch) Then
         '''If item was requested to be resubmitted for modification
         If bForceModify Then
                                                                                    lLine = 330
            '''Reload the same item
            objOrdSetItem.pbSelected = True
                                                                                    lLine = 340
            objOrdSetItem.pbModify = True
         Else
                                                                                    lLine = 360
            '''Load next selected item if there are any left
            sKey = ""
         End If
         '''If no more found or error - cancel batch
         mbCancelBatch = Not M_bBatchLoadNextItem(sErrMsg, sKey)
      End If
                                                                                    lLine = 400
   Else
                                                                                    lLine = 410
      mbCancelBatch = True  '''Just to make sure
   End If
   If mbCancelBatch Then
                                                                                    lLine = 420
      Call M_bBatchEnd
   End If
                                                                                    lLine = 500
   M_bBatchSubmitItem = sErrMsg = ""
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   M_bBatchSubmitItem = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   Call M_bBatchEnd
   
End Function

Private Function M_bBatchEnd() As Boolean
'''************************************************************************
'''Called By:     M_bBatchSubmitItem
'''               P_bSubmitOrders
'''               P_OrderStatus
'''               mobjSubmitTmr_pLibTimerEvent
'''               mobjStatusBox_pevLibStatusCancel
'''Arguments:
'''
'''Description:   Cleanup after the batch
'''
'''Note:          see also M_bBatchStart
'''Warning:
'''Returns:       True  - if successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchEnd"

   Dim lLine            As Long
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd

   bResult = True
                                                                                                lLine = 10
   mbProcessingBatch = False
                                                                                                lLine = 30
   '''Make sure all events are removed and timer disabled
   Call mobjSubmitTmr.P_bEventsClearAll
                                                                                                lLine = 60
   If Not mobjStatusBox Is Nothing Then
                                                                                                lLine = 80
      Call M_bBatchUpdateStatusReport
                                                                                                lLine = 100
      '''If error information
      If Not mobjStatusBox.psMsgReport = "" Then
                                                                                                lLine = 120
         '''Allow user to see it
         Call LibApp_Wait(3, True)
      End If
                                                                                                lLine = 140
      '''Take down  and release status form
      mobjStatusBox.P_Display (False)
                                                                                                lLine = 160
      Set mobjStatusBox = Nothing
   End If
                                                                                                lLine = 180
   If Not mobjOESMan Is Nothing Then
                                                                                                lLine = 182
      Call mobjOESMan.RequestToUpdateOrders
   End If
                                                                                                lLine = 200
   Call PopulateOrdersGrid
                                                                                                lLine = 220
   '''If batch was initiated when user clicked Submit button of the Order sets form
   '''(order Sets tab was active )
   If mbBatchInitiatedLocally Then
                                                                                                lLine = 240
      '''Return to initial tab
      If Not (gobjAnPParent Is Nothing) Then Call gobjAnPParent.changetab(mlANP_TAB_ORDER_SETS)
                                                                                                lLine = 260
      mbBatchInitiatedLocally = False  '''Set to true in mcmdSubmit_Click
   End If
                                                                                                lLine = 280
   If Not (gobjAnPParent Is Nothing) Then Call gobjAnPParent.UnlockAnP(msMODULE_NAME)
                                                                                                lLine = 300
   Call cwobjCWComm.Message _
   (cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION)
                                                                                                lLine = 320
   'SCR 23627  Phil Crowder 8/13/02
   cwobjCWComm.Message cwiORDER_SETS_PROCESSING, "N", NCID_ASSESMENT_AND_PLAN_APPLICATION, NCID_ORDER_SETS_APPLICATION
                                                                                                                                                                  
   M_bBatchEnd = bResult
   
Exit Function
ErrHnd:
   bResult = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   Resume Next
   
End Function

Private Function M_bBatchUpdateStatus _
      (ByVal sOrdName As String, ByVal bAllowModify As Boolean, ByVal sOrdTypeName As String) As Boolean
'''************************************************************************
'''Called By:     M_bBatchSubmitItem
'''
'''Arguments:
''' sOrdName      Name of the Item
''' bAllowModify  Indicates whether Item need to submited with option to modify
'''
'''Description:   Update status of the batch
'''
'''Note:
'''Warning:
'''
'''Returns:       True  - if successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchUpdateStatus"
   
   Dim sPrompt          As String

On Error GoTo ErrHnd
                                                                                                
10    Call M_bBatchUpdateStatusReport

20    If bAllowModify Then
                                                                                                
30       sPrompt = LibStr_sBuildSentence(gsPROCESSING, gsITEM & gs_S_, gsSELECTED, gsFOR, gsMODIFICATION, "...")
                                                                                                
40       If gobjShared Is Nothing Then Set gobjShared = New CWShared
50       If gobjShared.IsAppMode(modeTheater) Then
            'SCR 49504 - Use different message text Theater since button caption is "Save" not "Submit"
60          mobjStatusBox.psMsgWarning = LibStr_sBuildSentence(gsPLEASE, gsPRESS, gsSAVE, gsTO, gsACCEPT, gsAN, gsORDER, gsOR, gsCLEAR, gsTO, gsCANCEL, gsIT)
70       Else
80          mobjStatusBox.psMsgWarning = LibStr_sBuildSentence(gsPLEASE, gsPRESS, gsSUBMIT, gsTO, gsACCEPT, gsAN, gsORDER, gsOR, gsCLEAR, gsTO, gsCANCEL, gsIT)
90       End If
      
100   Else
110      sPrompt = LibStr_sBuildSentence(gsPROCESSING, gsSELECTED, gsITEM & gs_S_) & vbCrLf & LibStr_sBuildSentence(gsCREATING, gsNEW, "... ")
120      mobjStatusBox.psMsgWarning = ""
130   End If

140   sPrompt = sPrompt & vbCrLf & UCase$(LibStr_sBuildSentence(gsPRESS, gsCANCEL, gsTO, gsSTOP))
   
150   With mobjStatusBox
160      .psMsgPrompt = sPrompt
170      .plValue = mobjStatusBox.plValue + 1
180      .psMsgCount = LibStr_sBuildSentence(sOrdTypeName, gsORDER, ": ", sOrdName)
190   End With

200   M_bBatchUpdateStatus = True
   
   Exit Function

ErrHnd:
   M_bBatchUpdateStatus = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Private Function M_bGetCount _
      (ByRef lSelected As Long, ByRef lModify As Long, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     OrderSets.Mtd_bPushOrder
'''
'''Arguments:
''' lSelected     Count of all selected  to be returned byref
''' lModify       Count of selected for modification to be returned byref
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Count number of all items that was selected for submition
'''               and number of items that was selected for modifications
'''
'''Note:          Number of selected without modification can be calculated as a
'''               difference between all selected and selected for modification
'''Warning:
'''
'''Returns:       True
'''                  At least one selected item was found and no errors
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bGetCount"

   Dim bResult          As Boolean
   Dim objOrdSetItem    As clsOrderSetItem
   
On Error GoTo ErrHnd

   lSelected = 0
   lModify = 0
   If Not mcolOrdSets Is Nothing Then
      For Each objOrdSetItem In mcolOrdSets
         If objOrdSetItem.pbSelected Then
            lSelected = lSelected + 1
            '''Current rule is: item can not be set to modify if it is not selected
            If objOrdSetItem.pbModify Then
               lModify = lModify + 1
            End If
         ElseIf objOrdSetItem.pbModify Then
            Call LibApp_Error _
            (1, LibStr_sBuildSentence(gsINVALID, gsMODIFICATION, gsREQUEST), _
             msMODULE_NAME, sPROC_NAME, objOrdSetItem.psName _
            )
         End If
      Next objOrdSetItem
   End If
      
   M_bGetCount = lSelected > 0
   
Exit Function
ErrHnd:
   M_bGetCount = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Private Sub InitOE()

    If mobjOEGetSMan Is Nothing Then
        Set mobjOEGetSMan = New GEMS_OE.OE_GetSessionManager
    End If
  
    If mobjOESMan Is Nothing Then
        Set mobjOESMan = mobjOEGetSMan.GetSessionManager
    End If
  
End Sub

Private Function M_bBatchStart(ByVal lCount As Long, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     P_bSubmitOrders
'''
'''Arguments:
''' sErrMsg       Error information if any - to be returned byref
''' lCount        Indicates number of Items in the batch
'''
'''Description:   Make sure all required entities are initialized
'''
'''Note:          See also M_bBatchEnd
'''
'''Warning:
'''
'''Returns:       True  - if successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bBatchStart"

   Dim lLine            As Long
   
On Error GoTo ErrHnd

   sErrMsg = ""
                                                                                          lLine = 10
   msOrderStatus = ""
                                                                                          lLine = 12
   mbCancelBatch = False
                                                                                          lLine = 14
   Call cwobjCWComm.Message _
   (cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION)
                                                                                             lLine = 40
                                                                                             
   'SCR 23627  Phil Crowder 8/13/02
   cwobjCWComm.Message cwiORDER_SETS_PROCESSING, "Y", NCID_ASSESMENT_AND_PLAN_APPLICATION, NCID_ORDER_SETS_APPLICATION
   
   InitOE
                                                                                          lLine = 120
   If mobjStatusBox Is Nothing Then
                                                                                          lLine = 130
      Set mobjStatusBox = New mclsLibStatusBox
   End If
                                                                                          lLine = 140
   With mobjStatusBox
                                                                                          lLine = 160
      .plMaxVal = lCount
                                                                                          lLine = 180
      .pbAllowCancel = True
                                                                                          lLine = 200
      .pbDisplayCount = True
                                                                                          lLine = 220
      .pbProgressBar = True
                                                                                          lLine = 240
      .psCaption = gsORDER_SETS
                                                                                          lLine = 260
      .psMsgPrompt = ""
                                                                                          lLine = 270
      .psMsgCount = ""
                                                                                          lLine = 280
      .psMsgReport = ""
                                                                                          lLine = 290
      .psMsgWarning = ""
      
      Call .P_Display(True, True)
                                                                                          lLine = 300
      '''Place the box in the left top corner to avoid competing with other boxes
      Call .P_Move(0, 0, mlSTATUS_WIDTH)
   End With
                                                                                          lLine = 400
   M_bBatchStart = sErrMsg = ""
   
Exit Function
ErrHnd:
   M_bBatchStart = False
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE & " " & lLine)
   
End Function

Public Function P_bAddOrdItem _
   (objItem As Object, ByVal sKey As String, ByVal bQueue As Boolean, _
    ByRef sErrMsg As String _
   ) As Boolean
'''************************************************************************
'''Called By:     OrderSets.Mtd_bPushOrder
'''
'''Arguments:
''' objItem       Order Item(type Rx, Lab or Med) requested to be added to the queue
''' sKey          Key of the specified queue Item - used for Mtd_bRemoveOrder
''' bQueue        Indicates whether Item is part of the queue (else - template item)
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Add specified Item to collection of order items and if form is loaded
'''               Add It to the grid as well
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bAddOrdItem"

   Dim lLine            As Long
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   If M_bAddOrdItemToCollection(objItem, sKey, bQueue, sErrMsg) Then
                                                                                          lLine = 10
      If mbFormIsLoaded Then
                                                                                          lLine = 40
         bResult = AddOrdItemToGrid(sKey, True, sErrMsg)
      Else
         bResult = True
      End If
   End If
   
   P_bAddOrdItem = bResult
   
Exit Function
ErrHnd:
   P_bAddOrdItem = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function

Public Function P_bClearQueue _
      (Optional ByVal bMed As Boolean = True, _
       Optional ByVal bLab As Boolean = True, _
       Optional ByVal bRad As Boolean = True _
      ) As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''Arguments:
''' bMed          Request to remove all queued Med Items
''' bLab          Request to remove all queued Lab Items
''' bRad          Request to remove all queued Rad Items
'''
'''Description:   Remove all requested items from the queue and if mainform is loaded
'''               from the screen
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bClearQueue"

   Dim lLine            As Long
   Dim objOrdSetItem    As clsOrderSetItem
   Dim bRemove          As Boolean
   
On Error GoTo ErrHnd
   
   For Each objOrdSetItem In mcolOrdSets
      If objOrdSetItem.pbQueueItem Then
         Select Case objOrdSetItem.plType
            Case itMed
               bRemove = bMed
            Case itLab
               bRemove = bLab
            Case itRad
               bRemove = bRad
            Case Else
               bRemove = False
         End Select
         If bRemove Then
            Call mcolOrdSets.Remove(objOrdSetItem.psKey)
         End If
      End If
   Next objOrdSetItem
   
   P_bClearQueue = PopulateOrdersGrid
   
Exit Function
ErrHnd:
   P_bClearQueue = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
      
End Function

Public Function P_bOpen(ByVal lParenthWnd As Long, ByRef lhWnd As Long) As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' lParenthWnd   Handle to the parent form
''' lhWnd         handle to the Order Sets form to be return byref
'''
'''Description:   Open view session
'''               Initialize entities required for the view session
'''               Load form and data
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bOpen"

   Dim lLine            As Long
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   '''If was able to read and load template Items
   If gbGreenList = False Then
        If P_bLoadTemplateItems Then
           '''If was able to initialize and load the form and set references to required controls
           If InitGUI Then
                                                                                                  lLine = 10
              '''Attach order sets form to specified parent form
              Call SetParent(mfrmOrdSets.hwnd, lParenthWnd)
                                                                                                  lLine = 20
              lhWnd = mfrmOrdSets.hwnd
                                                                                                  lLine = 30
              '''If was able to populate the grid with all items
              If PopulateOrdersGrid Then
                 bResult = True
              End If
           End If
        End If
   Else
           If InitGUI Then
                                                                                                  lLine = 10
              '''Attach order sets form to specified parent form
              Call SetParent(mfrmOrdSets.hwnd, lParenthWnd)
                                                                                                  lLine = 20
              lhWnd = mfrmOrdSets.hwnd
                                                                                                  lLine = 30
              '''If was able to populate the grid with all items
              If PopulateOrdersGrid Then
                 PopulateStandaloneGrid
                 bResult = True
              End If
           End If
   
   End If
   
   P_bOpen = bResult
   
Exit Function
ErrHnd:
   P_bOpen = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Function
Public Function P_bClose() As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments
'''
'''Description:   Close view session - do necessary housekeeping
'''Note:          Release entities required for the view session
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bClose"
   
   Dim lLine            As Long
   
On Error GoTo ErrHnd
   
   'SCR 27530  Phil Crowder 9/11/02
   Call mobjSubmitTmr.P_bEventsClearAll
                                                                                                lLine = 10
   Set mcolOrdSets = New Collection
                                                                                                lLine = 20
   If Not mfrmOrdSets Is Nothing Then
                                                                                                lLine = 30
      If mbFormIsLoaded Then
                                                                                                lLine = 40
         Call Unload(mfrmOrdSets)
      End If
                                                                                                lLine = 50
      Set mfrmOrdSets = Nothing
   End If
   
   mbFormIsLoaded = False
      
   P_bClose = True
   
Exit Function
ErrHnd:
   P_bClose = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   
End Function
Public Function P_bRemoveOrdItem(ByVal sKey As String, Optional sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     P_bLoadTemplateItems
'''
'''Arguments:
''' sKey          Key of the item requested to be removed
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Remove specified Item
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bRemoveOrdItem"

   Dim lLine            As Long
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
                                                                                          lLine = 10
   Call mcolOrdSets.Remove(sKey)
                                                                                          lLine = 20
   If mbFormIsLoaded Then
                                                                                          lLine = 30
      bResult = PopulateOrdersGrid
   Else
      bResult = True
   End If
                                                                                          lLine = 40
   
   P_bRemoveOrdItem = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   P_bRemoveOrdItem = False
   Call LibApp_Error _
   (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, _
    LibStr_sBuildSentence(gsLINE, lLine & ", ", gsKEY, sKey) _
   )
End Function

Public Function P_bSubmitOrders(ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     mcmdSubmit_Click
'''               OrderSets.Mtd_bSubmitAll
'''Arguments:
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Submit all selected items to OE for processing
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bSubmitOrders"

   Dim lLine            As Long
   Dim lSelectedCount   As Long
   Dim objMsPnt         As New mclsLibMsPnt
   
On Error GoTo ErrHnd
                                                                                             lLine = 10
   sErrMsg = ""
                                                                                             lLine = 20
   Call objMsPnt.Mtd_SetMousePointer(Screen)
                                                                                             lLine = 40
   '''If there are any items that was requested to be processed
   If M_bGetCount(lSelectedCount, mlModifyCount, sErrMsg) Then
                                                                                             lLine = 80
      '''If was able to initiate batch
      If M_bBatchStart(lSelectedCount, sErrMsg) Then
                                                                                             lLine = 100
         mbProcessingBatch = M_bBatchLoadNextItem(sErrMsg)
      End If
   Else
                                                                                             lLine = 110
      sErrMsg = LibStr_sBuildSentence(gsNO, gsSELECTED, gsITEM & "s", gsFOUND)
   End If
                                                                                             lLine = 120
   If Not mbProcessingBatch Then
                                                                                             lLine = 200
      Call pobjMsgBox.LibMsg_lOKCancel(lMSGOK_NONE, , , "", sErrMsg, gsORDER_SETS)
                                                                                             lLine = 220
      Call M_bBatchEnd
   End If
                                                                                             lLine = 300
   P_bSubmitOrders = sErrMsg = ""
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   P_bSubmitOrders = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   Call M_bBatchEnd
   
End Function

Public Sub P_OrderStatus(ByVal sMsg As String)
'''************************************************************************
'''Called By:     OrderSets.Message
'''
'''Arguments:
''' sMsg          Contains "data" assoicated with current order
'''
'''Description:   Based on the current message update status of the last order submitted
'''               and perform appropriate action: Initiate next submission or end batch
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_OrderStatus"

   Dim lLine            As Long
   Dim bUnknownMessage  As Boolean  '''<DSN
   
On Error GoTo ErrHnd

   '''If waiting for submission result from OE
   If mbProcessingBatch Then
                                                                  lLine = 20
      If sMsg = "RECEIEVED ORDER" Then
                                                                  lLine = 40
         msOrderStatus = ""   '''Successfully submitted
      ElseIf Left(UCase(sMsg), 6) = "HL7ERR" Then
                                                                  lLine = 60
         '''If order was canceled by user
         If Right(sMsg, 3) = "803" Then
                                                                  lLine = 80
            msOrderStatus = gsCANCELED
         Else
                                                                  lLine = 100
            msOrderStatus = gsFAILED
         End If
      Else
         bUnknownMessage = True
      End If
                                                                  lLine = 110
      '''If one of the order sets specific messages
      If Not bUnknownMessage Then
                                                                  lLine = 120
         '''If failed to Initiate submission of the next order (all processed or error)
         If Not M_bBatchLoadNextItem("") Then
                                                                  lLine = 130
            Call M_bBatchUpdateStatusReport
                                                                  lLine = 140
            Call M_bBatchEnd
         End If
      End If
   End If

Exit Sub
ErrHnd:
   Call LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine & vbCrLf & sMsg)

End Sub

Public Sub P_ResizeView _
   (ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
    Optional ByRef sErrMsg As String _
   )
'''************************************************************************
'''Called By:     OrderSets.ResizeView
'''
'''Arguments:
''' lLeft         Contains the "left" coordinate for the OLE server.
''' lTop          Contains the "top" coordinate for the OLE server.
''' lWidth        Contains the "width" for the OLE server.
''' lHeight       Contains the "height" for the OLE server.
''' sErrMsg       Optional error information if any - to bew returend byref
'''
'''Description:   Position the form on the screen
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
   Const sPROC_NAME     As String = "P_ResizeView"

   Dim objMousePointer  As New mclsLibMsPnt
   Dim lLine            As Long
   
On Error GoTo ErrHnd
   
   If mbFormIsLoaded Then
                                                                                          lLine = 10
      Call objMousePointer.Mtd_SetMousePointer(Screen)
                                                                                          lLine = 20
      With mfrmOrdSets
                                                                                          lLine = 30
         '''Avoid screen flickering
         .Visible = False
                                                                                          lLine = 40
         '''  Make sure that "WindowState" is "Normal", so the form can be resized.  This is
         '''    needed because a window CANNOT be moved or resized if it is "minimized" or
         '''    "maximized".
         .WindowState = vbNormal
                                                                                          lLine = 50
         '''Resize form to fill entire "client area"
         ''''''NOTE: This generates a "resize" event
                                                                                          lLine = 60
         .Move lLeft, lTop, lWidth, lHeight
         
         If gbGreenList = False Then
            Call .Show
         Else
            If P_bLoad_Template_Favs Then
               Call .Show
            End If
         End If
         'jrm:  Let all events catch up before going visible
         DoEvents
         
         .Visible = True
                                                                                          lLine = 90
      End With
   Else
      sErrMsg = LibStr_sBuildSentence(gsVIEW, gsIS, gsNOT, gsOPEN)
      Call LibApp_Error(1, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   End If
   
Exit Sub
ErrHnd:
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)

End Sub


Private Sub Class_Initialize()
    On Error Resume Next
   Set mcolOrdSets = New Collection
   Set mobjSubmitTmr = New mclsLibTimer
   
End Sub

Private Sub Class_Terminate()
On Error Resume Next


   '''Batch control objects
   Set mcolOrdSets = Nothing
   Set mobjStatusBox = Nothing
   Set mobjSubmitTmr = Nothing
   '''GUI interface
   Set mcmdDeSelectAll = Nothing
   Set mcmdSelectAll = Nothing
   Set mcmdRefreshData = Nothing
   Set mcmdSubmit = Nothing
   Set mvsgOrderItems = Nothing
   Set mvsgStandalone = Nothing
   Set mfrmOrdSets = Nothing
   '''OE Interface
   Set mobjOEGetSMan = Nothing
   Set mobjOESMan = Nothing

End Sub


Private Sub mcmdDeSelectAll_Click()
   Call M_bSelectAll(False)
End Sub

Private Sub mcmdRefreshData_Click()
   If P_bLoadTemplateItems Then
      Call PopulateOrdersGrid
      Call PopulateStandaloneGrid
   End If
End Sub

Private Sub mcmdSaveOrders_Click()
  'SCR 20107  3/13/03 Phil Crowder
  'JT SCR 36572 - Updated code to check IsAppMode
  If Len(gobjShared.CmdLineSwitches.Item("GREEN")) > 0 Or gobjShared.IsAppMode(modeCHCSI_GUI) Then
    Call SaveOrders
  End If
End Sub

Private Sub mcmdSelectAll_Click()
   Call M_bSelectAll(True)
End Sub

Private Sub mcmdSubmit_Click()
   mbBatchInitiatedLocally = True
   Call P_bSubmitOrders("")
End Sub

Public Sub SaveOrders()
    
    Dim oEncParent As Object
    Dim objTemplateOps   As CHCSII_TemplateClient.TemplateOps
    Dim objTemplate      As CHCSII_TemplateClient.Template
    Dim objOrdSetItem As clsOrderSetItem
    Dim colOrders As Collection
    Dim sName As String
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    
    If oEncParent Is Nothing Then Set oEncParent = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    If Not oEncParent Is Nothing Then
       If Not oEncParent.currentencounter Is Nothing Then
            sName = "Orders from encounter " & CStr(oEncParent.currentencounter.encounterid)
       End If
    End If
    
    Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
    Call objTemplateOps.Initialize(cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser)
    Set objTemplate = objTemplateOps.NewTemplate(EncounterTemplateType)
    Set colOrders = New Collection
    
    For Each objOrdSetItem In mcolOrdSets
         If objOrdSetItem.pbSelected Then
            colOrders.Add objOrdSetItem.ExportItem
         End If
    Next
    
    objTemplate.OrderList.Import SerializeCollection(colOrders, gsListDelim)
    If colOrders.Count = 0 Then
       MsgBox "No order selected to add to a template.", vbOKOnly + vbInformation
    Else
       objTemplate.TemplateName = sName
       Screen.MousePointer = vbDefault
       
       'JT BEGIN SCR 39281 - Need to notify modules that modal dialog is displaying
       Call cwobjCWComm.Message(cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION)
           
       If Not objTemplate.SaveAs Then
           MsgBox "Template was not saved.", vbOKOnly + vbInformation
       End If
       
       Call cwobjCWComm.Message(cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION)
       'JT END SCR 39281
       
    End If

    GoTo CleanUpAndExit
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "SaveOrders()", "OrderSets", vbExclamation)
    
CleanUpAndExit:
    Set objOrdSetItem = Nothing
    Set objTemplate = Nothing
    Set objTemplateOps = Nothing
    Set oEncParent = Nothing
    Screen.MousePointer = vbDefault

End Sub


Private Sub mobjStatusBox_pevLibStatusCancel()
'''************************************************************************
'''Called By:     When User cancels batch
'''
'''Arguments:
'''
'''Description:   Cancel Batch
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If successful
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "mobjStatusBox_pevLibStatusCancel"
   
On Error GoTo ErrHnd
   
   mobjStatusBox.psMsgPrompt = UCase$(LibStr_sBuildSentence(gsPROCESSING, gsCANCEL, gsREQUEST, "..."))
   mbCancelBatch = True
   '''Just to make sure - In the case if batch hangs -
   '''load terminate batch event wit 10 seconds delay to allow to complete
   '''verify or create new if already started
   Call mobjSubmitTmr.P_sAddEventRegularInterval("s", 10, "", gsCANCEL & gsBATCH)
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   Call M_bBatchEnd
   
End Sub

Private Sub mobjSubmitTmr_pLibTimerEvent _
      (ByVal tEventDateTime As Date, ByVal dDelay As Double, ByVal sMsg As String, _
       ByVal sID As String, ByVal lCallNo As Long, ByVal vTag As Variant _
      )
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sID           Key of the Item
''' sMsg          Indicates whether Item Need to be submitted for modification or
'''               without option to modify
'''
'''Description:   Allows asynchronous call to OE to avoid reentry in the message routine.
'''
'''Note:          OE can change request from Create New to Update Screen (Modify) -
'''               which will require user interaction.
'''               Therefore we can not just stay in the loop and keep on pushing the items.
'''               Order sets must wait for message from OE indicating result
'''               of the submission: success, user cancel, error...
'''               And - Message can be called before submit call to OE ends (create new).
'''
'''
'''               Also refer to M_bBatchLoadNextItem
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "mobjSubmitTmr_pLibTimerEvent"

   Dim lLine            As Long
   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd
                                                                                                lLine = 10
   '''If not finished processing previous item
   If Not bIn_static Then
                                                                                                lLine = 20
      bIn_static = True
      '''If ready to process - remove event to make sure it does not fire again
      Call mobjSubmitTmr.P_bEventRemove(sID)
                                                                                                lLine = 40
      If sID = gsCANCEL & gsBATCH Then
                                                                                                lLine = 60
         Call M_bBatchEnd
      Else
                                                                                                lLine = 80
         Call M_bBatchSubmitItem(sID, sMsg = msOPT_MODIFY, "")
      End If
                                                                                                lLine = 100
      bIn_static = False
   End If
                                                                                                lLine = 120
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   bIn_static = False
   Call M_bBatchEnd
   
End Sub


Private Sub mvsgOrderItems_AfterSort(ByVal Col As Long, Order As Integer)
   Static bIn_static As Boolean
   
On Error Resume Next

   If Not bIn_static Then
      bIn_static = True    '''Just to make sure
      With mvsgOrderItems
         '''If user attempts to sort selected column
         If Col = mlCOL_SELECT Then
            '''Use default sort and
            Call .Select(0, mlCOL_SORT)
            '''Toggle an order
            If .ColSort(mlCOL_SORT) = flexSortStringNoCaseDescending Then
               .ColSort(mlCOL_SORT) = flexSortStringNoCaseAscending
            Else
               .ColSort(mlCOL_SORT) = flexSortStringNoCaseDescending
            End If
            .Sort = flexSortUseColSort
         End If
      End With
      bIn_static = False
   End If
   
End Sub

Private Sub mvsgOrderItems_Click()
'''************************************************************************
'''Called By:     OrderSets.ResizeView
'''
'''Arguments:
'''
'''Description:   Update selected and modified properties
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
   Const sPROC_NAME     As String = "mvsgOrderItems_Click"

   Dim lLine            As Long
   Dim bVal             As Boolean
   Dim objOrdSetItem    As clsOrderSetItem
   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd
   
   If Not bIn_static Then
                                                                                    lLine = 10
      bIn_static = True
                                                                                    lLine = 20
      With mvsgOrderItems
         
         '''If not outside of the selectable region and colums that can be modified by the user
         If .MouseRow > 0 And .row > 0 And (.Col = mlCOL_SELECT Or .Col = mlCOL_MODIFY) Then
                                                                                    lLine = 30
            Set objOrdSetItem = mcolOrdSets(.TextMatrix(.row, mlCOL_KEY))
                                                                                    lLine = 32
            '''<SCR 254(Pharmacy) - must use RxPOE interface for BDM pharmacy
            If (objOrdSetItem.plHLOrderType = hlRX_ORDERS_CODE) And (gobjShared.pbMTFSupportsBDM) Then
                                                                                    lLine = 34
               Call pobjMsgBox.LibMsg_lOKCancel(lMSGOK_CURRENTLY_NOT_SUPPORTED, , , "Select Rx Order", "by RxTFC", gsORDER_SETS)
                                                                                    lLine = 36
               Call PopulateOrdersGrid
            '''>SCR 254
            Else
                                                                                    lLine = 40
               bVal = .TextMatrix(.row, .Col)
                                                                                    lLine = 50
               bVal = Not bVal
                                                                                    lLine = 60
               Select Case .Col
                  Case mlCOL_SELECT
                                                                                    lLine = 70
                     .TextMatrix(.row, .Col) = bVal
                                                                                    lLine = 80
                     objOrdSetItem.pbSelected = bVal
                                                                                    lLine = 90
                     If Not bVal Then
                                                                                    lLine = 100
                        .TextMatrix(.row, mlCOL_MODIFY) = False
                                                                                    lLine = 110
                        objOrdSetItem.pbModify = False
                                                                                    lLine = 120
                     End If
                  Case mlCOL_MODIFY
                                                                                    lLine = 200
                     .TextMatrix(.row, .Col) = bVal
                                                                                    lLine = 210
                     objOrdSetItem.pbModify = bVal
                                                                                    lLine = 220
                     If bVal Then
                                                                                    lLine = 230
                        .TextMatrix(.row, mlCOL_SELECT) = True
                                                                                    lLine = 240
                        objOrdSetItem.pbSelected = True
                                                                                    lLine = 250
                     End If
               End Select
            End If
                                                                                    lLine = 260
            Set objOrdSetItem = Nothing
         End If
      End With
      bIn_static = False
   End If
   
Exit Sub
ErrHnd:
   Set objOrdSetItem = Nothing
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & lLine)
   bIn_static = False
End Sub
Public Function P_bLoad_Template_Favs() As Boolean


   Dim objTemplateOps  As CHCSII_TemplateClient.TemplateOps
   
   Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
   objTemplateOps.Initialize cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser
   
   Set mfrmOrdSets.ucTemplateNavigation1.TemplateOps = objTemplateOps
   Set mfrmOrdSets.ucTemplateNavigation1.TemplateHistoryOps = objTemplateOps.TemplateHxOps
   
  'jrm:add only for encounters
    mfrmOrdSets.ucTemplateNavigation1.RunningMode = eEncounterTemplatesOnly
    
  'jrm:Add favorites to templatehx
  mfrmOrdSets.ucTemplateNavigation1.TemplateHistoryOps.Add_Favorites objTemplateOps.GetFavoriteTemplates(False, True)
   
   P_bLoad_Template_Favs = True
End Function

Private Sub mvsgOrderItems_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  'SCR 20107  3/13/03 Phil Crowder
  'JT SCR 36572 - Updated code to check IsAppMode
  If Len(gobjShared.CmdLineSwitches.Item("GREEN")) > 0 Or gobjShared.IsAppMode(modeCHCSI_GUI) Then
    If Button = vbRightButton Then
        mfrmOrdSets.ShowPopupMenu
    End If
  End If
End Sub


Public Function P_bPersistHistory() As Boolean
mfrmOrdSets.ucTemplateNavigation1.PersistHistory
P_bPersistHistory = True
End Function
Public Function P_bRetrieveHistory() As Boolean
'SCR 33723 3/20/03  Phil Crowder
'1. Order Queued Orders
'2. Orders Sets module Inited by AnP
'3. Select OrderSets tab
'4. AnP sends ForeGround message (that fires this function) because module has been inited.
'Need One Error Resume Next for this case because Nav control History has not been loaded
'yet in the this case.
On Error Resume Next

mfrmOrdSets.ucTemplateNavigation1.RetrieveHistory
P_bRetrieveHistory = True
End Function

Private Sub mvsgStandalone_GotFocus()

   On Error GoTo Error_Handler

   If mvsgStandalone.row > 0 Then
      TemplateSelected mvsgStandalone.TextMatrix(mvsgStandalone.row, sgcKEY), False
   End If
      
Error_Handler:
   If CBool(Err) Then Call HandleErrors(Err, msMODULE_NAME, "mvsgStandalone_GotFocus")
   
End Sub

Private Sub mvsgStandalone_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo Error_Handler

Static lPrevRow As Long
Static lPrevCol As Long
   
Dim lNewRow As Long
Dim lNewCol As Long
   
   ' get coordinates
   lNewRow = mvsgStandalone.MouseRow
   lNewCol = mvsgStandalone.MouseCol
    
   ' update tooltip text
   If lPrevCol <> lNewCol Or lPrevRow <> lNewRow Then
       lPrevRow = lNewRow
       lPrevCol = lNewCol
       If lNewRow > 0 And lNewCol >= 0 Then
         mvsgStandalone.ToolTipText = mvsgStandalone.TextMatrix(lNewRow, sgcFULL_NAME)
       End If
   End If
    
Error_Handler:
   'Ignore Error - we don't care about an error in this routine
   
End Sub

Private Sub mvsgStandalone_SelChange()
   
   On Error GoTo Error_Handler

   If mvsgStandalone.row > 0 Then
      TemplateSelected mvsgStandalone.TextMatrix(mvsgStandalone.row, sgcKEY), False
   End If
      
Error_Handler:
   If CBool(Err) Then Call HandleErrors(Err, msMODULE_NAME, "mvsgStandalone_SelChange")
   
End Sub

Private Sub ucNavigation_GotFocus()

   On Error GoTo Error_Handler

   TemplateSelected mfrmOrdSets.ucTemplateNavigation1.TemplateHistoryOps.CurrentTemplateHxItem.TemplateItem.TemplateID, False
   If mbY2TDisplay Then mvsgStandalone.row = -1
   
Error_Handler:
   If CBool(Err) Then Call HandleErrors(Err, msMODULE_NAME, "ucNavigation_GotFocus")
   
End Sub

Private Sub ucNavigation_Navigate(rsTemplateHxID As String)

   On Error GoTo Error_Handler

   TemplateSelected mfrmOrdSets.ucTemplateNavigation1.TemplateHistoryOps.CurrentTemplateHxItem.TemplateItem.TemplateID, True
   If mbY2TDisplay Then mvsgStandalone.row = -1
   
Error_Handler:
   If CBool(Err) Then Call HandleErrors(Err, msMODULE_NAME, "ucNavigation_Navigate")
   
End Sub

Friend Property Get OrderSetsForm() As frmOrderSets
'This is a temporary method that will allow the ordersets class to
'access the form instance until this layering can be cleaned out
   Set OrderSetsForm = mfrmOrdSets
End Property

Friend Sub PopulateStandaloneGrid()
'SCR 52899 - Changed to use Version 6 of VSFlexGrid to avoid exception fault being caused by Version 7

On Error GoTo ErrHnd

Dim oTemplateOps As CHCSII_TemplateClient.TemplateOps
Dim oTemplate As CHCSII_TemplateClient.TemplateLite
Dim oTemplates As Collection
Dim sTruncatedName As String

   If Not mvsgStandalone Is Nothing And mbY2TDisplay Then
   
      Set oTemplateOps = New CHCSII_TemplateClient.TemplateOps
      Set oTemplates = oTemplateOps.DoStandaloneOrderSearchLite()
      
      mvsgStandalone.Redraw = False
      
      'Remove any items in the grid but not the column header row
      mvsgStandalone.Rows = 1
      
      'Populate the template items in the grid
      For Each oTemplate In oTemplates
         sTruncatedName = IIf(Len(oTemplate.TemplateName) <= 30, oTemplate.TemplateName, Left(oTemplate.TemplateName, 30) & " ...")
         mvsgStandalone.AddItem sTruncatedName & vbTab & oTemplate.TemplateID & vbTab & oTemplate.TemplateName
      Next
      
      mvsgStandalone.Redraw = True
       
   End If
   
   Exit Sub
   
ErrHnd:
   'Clean Up Objects
   Set oTemplateOps = Nothing
   Set oTemplates = Nothing
   Set oTemplate = Nothing
   
   If CBool(Err) Then Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, "PopulateStandaloneGrid", vbNullString)

End Sub
