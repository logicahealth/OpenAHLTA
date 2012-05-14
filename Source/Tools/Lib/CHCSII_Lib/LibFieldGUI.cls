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
Attribute VB_Name = "mclsLibFieldGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        clsLibFieldGUI/LibFieldGUI.cls
'''
'''Description:   Reference to user controls and data field definition object
'''
'''
'''Modules Required:
'''
'''Note:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME              As String = "mclsLibFieldGUI."

Private Enum mnumCT_ControlType
   lCT_Undefined
   lCT_Check
   lCT_Option
   lCT_Text
   lCT_Label
   lCT_List
   lCT_DTPicker
End Enum


''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mbEnabled             As Boolean  '''Indicates whether value can be changed by user
Private mbVisible             As Boolean
Private mbControlCaptionSet   As Boolean
Private mlControlTypeData     As mnumCT_ControlType
Private mlListKey             As Long
Private msFormat              As String
Private mbDescAsToolTip       As Boolean

'''Indicates whether to update properties of associated controls. Defaults to False
Public pbFreezeGUI            As Boolean
'''Indicates whether value displayed on the screen will be automatically updated
'''when Item Or Data field value is changed.
'''Note: This property is ignored if  value is applied directly to the GUI field pvValue property.
'''      In this case screen will always be updated automatically.
'''Same rules apply to Caption property
'''Defaults to True
Public pbAutoUpdate                       As Boolean
'''Indicates whether value assigned to GUI field pvValue property will be
'''automaticaly assigned to Item and Data field value. defaults to false
'''Also refer to P_AcceptControlValue
Public pbAutoAcceptControlValue           As Boolean

''''''''''''''' Records Section '''''''''''''''''''''''''''''''''''''''''


''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mctlCaption           As Control
Private mctlData              As Control

Private mobjList              As mclsLibList

Private WithEvents mobjItem   As mclsLibItem
Attribute mobjItem.VB_VarHelpID = -1

''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

Public Event pevCtlDataSet()


Friend Property Set fobjItem(objItem As mclsLibItem)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W     Default
'''               R  Y(f)  New
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
   Const sPROC_NAME     As String = "Set fobjItem"

On Error GoTo ErrHnd

   Set mobjItem = objItem
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property


Private Sub M_ApplyControlProps()
'''************************************************************************
'''Called By:     Let pbEditable
'''               Set pctlData
'''               Let pbEnabled
'''Arguments:
'''
'''Description:   Make sure that property settings are reflected in the GUI
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
   Const sPROC_NAME     As String = "M_ApplyControlProps"

On Error GoTo ErrHnd

   If mbControlCaptionSet Then
100   mctlCaption.Enabled = mbEnabled
   End If
   If mlControlTypeData > lCT_Undefined Then
200   mctlData.Enabled = mbEnabled
      If mlControlTypeData = lCT_Text Then
220      mctlData.Locked = Not mobjItem.pbEditable
      End If
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
         (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Sub

Private Function M_bDisplayCaption() As Boolean
'''************************************************************************
'''Called By:     P_bRefresh
'''               Let psCaption
'''Arguments:
'''
'''Description:   Display current caption
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - True
'''               Else           - False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bDisplayCaption"

   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd

   If (Not bIn_static) And (Not pbFreezeGUI) Then
      bIn_static = True
      If mbControlCaptionSet Then
100      mctlCaption.Caption = mobjItem.psName
         ''If Item description was requested to be used as tool tip
         If mbDescAsToolTip Then
200         mctlCaption.ToolTipText = mobjItem.psDesc
         End If
      End If
      
      bIn_static = False
   End If
   
   M_bDisplayCaption = True
   
Exit Function
ErrHnd:
   M_bDisplayCaption = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, psCaption & ", " & gsLINE_ & Erl, True)
   bIn_static = False
   
End Function

Private Function M_bDisplayValue(ByVal vValue As Variant) As Boolean
'''************************************************************************
'''Called By:     P_bRefresh
'''               Let pvValue
'''Arguments:
''' vValue
'''   Value requested to be displayed
'''
'''Description:   Display specified value. if pbAutoAcceptControlValue is true
'''               assign it to item value
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - True
'''               Else           - False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bDisplayValue"

   Static bIn_static    As Boolean
   Dim sMsg             As String
   
On Error GoTo ErrHnd

20 If (Not bIn_static) And (Not pbFreezeGUI) Then
      bIn_static = True
      '''If control is used
200   If mlControlTypeData > lCT_Undefined Then
         
         Select Case mlControlTypeData
            Case lCT_Check
300            mctlData.Value = Abs(gobjLibValid.LibValid_bEvaluateBoolean(vValue))
            Case lCT_Option
400            mctlData.Value = gobjLibValid.LibValid_bEvaluateBoolean(vValue)
            Case lCT_Text
500            mctlData.Text = M_sFormatedValue(vValue)
            Case lCT_Label
600            mctlData.Caption = M_sFormatedValue(vValue)
            Case lCT_List
700            Call M_bSetItemByItemData(vValue)
            Case lCT_DTPicker
800            If IsDate(vValue) Then
820               mctlData.Value = vValue
               Else
840               mctlData.Value = mctlData.MinDate
               End If
            Case Else
2000           sMsg = gobjLibStr.LibStr_sBuildSentence _
                     (gsCONTROL, gsTYPE, TypeName(mctlData), gsIS, gsNOT, gsSUPPORTED)
2100           Call gobjLibApp.LibApp_Error(0, sMsg, msMODULE_NAME, sPROC_NAME, "", True, False, False)
         End Select
         '''If Item description was requested to be used as tool tip
         If mbDescAsToolTip Then
3000        mctlData.ToolTipText = mobjItem.psDesc
         End If
      End If
      bIn_static = False
   End If
   
999   M_bDisplayValue = True
   
Exit Function
ErrHnd:
   M_bDisplayValue = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, psCaption & ", " & gsLINE_ & Erl, True)
   bIn_static = False
   
End Function

Private Function M_sFormatedValue(ByVal vValue As Variant) As String
'''************************************************************************
'''Called By:     M_bDisplayValue
'''               Get psValueFormated
'''
'''Arguments:
'''vValue
''' Value requested to be returned in the current format
'''
'''Description:   Return curent value in current format
'''
'''
'''Note:
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if successful
'''                  curent value in current format
'''               else
'''                  Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_sFormatedValue"

   Dim sResult          As String
   
On Error GoTo ErrHnd
      
20    If Len(msFormat) > 0 Then
40       sResult = Format$(gobjLibStr.LibStr_sEmptyIfNull(vValue), msFormat)
      Else
100      sResult = gobjLibStr.LibStr_sEmptyIfNull(vValue)
      End If

999   M_sFormatedValue = sResult
   
Exit Function
ErrHnd:
   M_sFormatedValue = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Private Sub M_SynchList()
'''************************************************************************
'''Called By:     Set pctlData
'''
'''Arguments:
'''
'''Description:   Populate object list with current list control items
'''
'''Note:          pctlData must be Listbox or ComboBox
'''
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_SynchList"

   Dim colCurrLst       As New Collection
   Dim objItem          As mclsLibItem
   Dim lIdx             As Long
   Dim vCurVal          As Variant
   
On Error GoTo ErrHnd

10    With mctlData
         '''If new list control contains Items
20       If .ListCount > 0 Then
40          For lIdx = 0 To .ListCount - 1
60             Set objItem = New mclsLibItem
80             objItem.pvValue = .ItemData(lIdx)
100            objItem.psName = .List(lIdx)
120            Call colCurrLst.Add(objItem)
140            Set objItem = Nothing
            Next lIdx
         End If
         '''If previous list has items
200      If Not mobjList Is Nothing Then
210         vCurVal = mobjItem.pvValue '''Store which Item is current
            '''Collect all previous items priour to removing them from mobjList
220         For Each objItem In mobjList
240            Call colCurrLst.Add(objItem)
            Next objItem
         End If
         '''Reset list properties of this object
290      Call P_ListClear
         '''If any information was collected
300      If colCurrLst.Count > 0 Then
320         For Each objItem In colCurrLst
340             Call P_sListAddItem(objItem.psName, , objItem.pvValue, objItem.pvTag)
            Next objItem
400         Set objItem = Nothing
         End If
      End With
      
500   pvValue = vCurVal '''Restore current Item
      
600   Set colCurrLst = Nothing
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Sub


Private Function M_vGetCurrentItemData() As Variant
'''************************************************************************
'''Called By:     Get pvValue
'''
'''Arguments:
''' lVal
''' Item Data value
'''
'''Description:   Return Item data value of the currently selected item
'''
'''Note:          This procedure assumes that mctlData is set and list type
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - Item data value of the currently selected item
'''               Else           - Empty
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_vGetCurrentItemData"

   Dim vResult          As Variant
   
On Error GoTo ErrHnd
   
   With mctlData
      If .ListIndex > -1 Then
         vResult = mobjList.pobjItem(CStr(.ItemData(.ListIndex))).pvValue
      End If
   End With
   
   M_vGetCurrentItemData = vResult
   
Exit Function
ErrHnd:
   M_vGetCurrentItemData = Empty
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Function


Private Function M_bSetItemByItemData(ByVal vValue As Variant) As Boolean
'''************************************************************************
'''Called By:     Let pvValue
'''
'''Arguments:
''' vValue
'''   Item Data value of the Item requested to be set as current
'''
'''Description:   Select first List item which item data match specified value
'''
'''Note:          This procedure assumes that mctlData is set and is a list type
'''               If empty will be set to none selected (index -1) or first item
'''               with empty value
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bSetItemByItemData"

   Static bIn_static    As Boolean
   Dim lIdx             As Long
   Dim bFound           As Boolean
   
On Error GoTo ErrHnd
   
      If Not bIn_static Then
20       mctlData.ListIndex = P_lIndexByItemData(vValue)
         bIn_static = False
      End If
   
999   M_bSetItemByItemData = True
   
Exit Function
ErrHnd:
   bIn_static = False
   M_bSetItemByItemData = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Private Property Get mobjListItem(Optional ByVal vByItemData As Variant) As mclsLibItem
'''************************************************************************
'''Called By:     pobjListItemByValue
'''               pobjListItemCurrent
'''               psListItem
'''Arguments:
''' vByItemData
'''   Optional Identifier of requested Item.
'''   If not specified properties of the current item will be returned
'''
'''Description:   Return requested list item
'''
'''Note:          If List control is set the properties of the current item may not match
'''               current pobjItem.pvValue unless P_AcceptControlValue is called
'''               Also refer to Get pvValue.
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if successful
'''                  Requested list item object
'''               else
'''                  nothing
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get mobjListItem"

   Dim sKey             As String
   Dim objResult        As mclsLibItem
   
On Error GoTo ErrHnd
      
      '''If specific item is not requested
100   If IsMissing(vByItemData) Then
         '''If list control is set
200      If mlControlTypeData = lCT_List Then
            '''If Item is selected
300         If mctlData.ListIndex >= 0 Then
               '''Return Properties corresponding to the item selected in the list control
320            sKey = mctlData.ItemData(mctlData.ListIndex)
            End If
         Else
400         sKey = mobjList.P_sFindByValue(pvValue)
         End If
      Else
500      sKey = mobjList.P_sFindByValue(vByItemData)
      End If
      
700   If Len(sKey) > 0 Then
720      Set objResult = mobjList.pobjItem(sKey)
      End If

999   Set mobjListItem = objResult
      
Exit Property
ErrHnd:
   Set mobjListItem = Nothing
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)
       
End Property

Public Sub P_AcceptControlValue()
'''************************************************************************
'''Called By:     Let psFormat
'''               Let pvValue
'''               Get pvValue
'''Arguments:
'''
'''Description:   Accept current control value as current field value
'''
'''Note:
''' Also refer to pbAutoAcceptControlValue
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_AcceptControlValue"

   Dim bResult          As Boolean
   Static bIn_static    As Boolean
   Dim vControlValue    As Variant
   
On Error GoTo ErrHnd
   
20 If (Not mctlData Is Nothing) And (Not bIn_static) Then
      bIn_static = True
      vControlValue = pvValue
      '''If current control value was changed - Avoid triggering unnecessary event.
      If Not (gobjLibStr.LibStr_sEmptyIfNull(vControlValue) = _
              gobjLibStr.LibStr_sEmptyIfNull(mobjItem.pvValue) _
             ) _
      Then
         '''Update Item value
         mobjItem.pvValue = vControlValue
      End If
      bIn_static = False
   End If
   
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, False, , False)
   bIn_static = False

End Sub

Public Function P_bRefresh(Optional ByVal bDoEvents As Boolean = True) As Boolean
'''************************************************************************
'''Called By:     Let psFormat
'''               mobjItem_pevNameChanged
'''               mobjItem_pevValueChanged
'''Arguments:
''' bDoEvents
'''   Indicates whether to inforce refresh with DoEvents
'''   Defaults to true
'''
'''Description:   If Data control is Set
'''                  Display current value and caption
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bRefresh"

   Dim bResult          As Boolean
   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd
   
20 If (Not mctlData Is Nothing) And (Not bIn_static) And (Not pbFreezeGUI) Then
      bIn_static = True
100   Call M_bDisplayValue(mobjItem.pvValue)
      '''Refresh caption
120   Call M_bDisplayCaption
      
      If bDoEvents Then
500      DoEvents
      End If
      
      bResult = True
      
      bIn_static = False
   End If
   
   P_bRefresh = bResult
   
Exit Function
ErrHnd:
   P_bRefresh = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, False, , False)
   bIn_static = False
   
End Function

Public Function P_bSetFocus() As Boolean
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
'''
'''Description:   If Data control is available: Set, visible and enabled
'''                  Set focus
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bSetFocus"

   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   If Not mctlData Is Nothing Then
      If mctlData.Visible And mctlData.Enabled Then
         Call mctlData.SetFocus
         bResult = True
      End If
   End If
   
   P_bSetFocus = bResult
   
Exit Function
ErrHnd:
   P_bSetFocus = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", False, , False)

End Function

Public Sub P_ClearDisplay()
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:   Based on the current data control associated with the field
'''               clear displayed value
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
   Const sPROC_NAME     As String = "P_ClearDisplay"

   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd
   
20 If (Not mctlData Is Nothing) And (Not bIn_static) And (Not pbFreezeGUI) Then
      bIn_static = True
      Select Case mlControlTypeData
         Case lCT_Check
300         mctlData.Value = 0
         Case lCT_Option
400         mctlData.Value = 0
         Case lCT_Text
500         mctlData.Text = ""
         Case lCT_Label
600         mctlData.Caption = ""
         Case lCT_List
700         mctlData.ListIndex = -1
         Case lCT_DTPicker
820         mctlData.Value = mctlData.MinDate
      End Select
      
      bIn_static = False
   End If
   
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, False, , False)
   bIn_static = False

End Sub

Public Function P_lIndexByItemData(ByVal vItemData As Variant, Optional ByVal sKey As String) As Long
'''************************************************************************
'''Called By:     P_ListRemoveItem or Any
'''               M_bSetItemByItemData
'''Arguments:
''' vItemData
'''   Item data value of the requested Item
''' sKey
'''   Optional Key value of the requested Item. empty if not found
'''
'''Description:   Find Item which ItemData matches specified value
'''               If found return corresponding index
'''
'''
'''Note:          Returns first occurance
'''
'''Warning:          '''Assumed that Data Control is set and is List or combo box
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  Index of the matching item
'''               else
'''                  -1
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_lIndexByItemData"
   
   Dim lResult          As Long
   Dim lIdx             As Long
   
On Error GoTo ErrHnd
   
   lResult = -1
   sKey = ""
   
   With mctlData
      For lIdx = 0 To .ListCount - 1
         If mobjList.pobjItem(CStr(.ItemData(lIdx))).pvValue = vItemData Then
            lResult = lIdx
            sKey = CStr(.ItemData(lIdx))
            Exit For
         End If
      Next lIdx
   End With
   
   P_lIndexByItemData = lResult
   
Exit Function
ErrHnd:
   P_lIndexByItemData = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Function

Public Function P_bLoadList(objList As mclsLibList, Optional bClear As Boolean = True) As Boolean
'''************************************************************************
'''Called By:     M_SynchList
'''
'''Arguments:
''' objList
'''   List of LibItem objects requested to be loaded
''' bClear
'''   Request to clear GUI list prior to adding new items. Defaults to true
'''
'''Description:   Adds List Item
'''
'''Note:          pctlData must be Listbox or ComboBox
'''               Unlike ItemData of the standard VB list controls
'''               vItemData is a variant and can Accept any data type
'''
'''Warning:
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bLoadList"

   Dim objItem          As mclsLibItem
   
On Error GoTo ErrHnd
   
   
100   If bClear Then
120      Call P_ListClear
      End If
200   For Each objItem In objList
220      Call P_sListAddItem(objItem.psName, , objItem.pvValue, objItem.pvTag)
300      With mobjList.pobjItemNew
310         .plAllowComments = objItem.plAllowComments
320         .plMinLenComments = objItem.plMinLenComments
330         .psComments = objItem.psComments
340         .psDesc = objItem.psDesc
350         .pvValueDefault = objItem.pvValueDefault
         End With
      Next objItem
      
999   P_bLoadList = True
   
Exit Function
ErrHnd:
   P_bLoadList = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Function P_sListAddItem _
      (ByVal sItem As String, Optional ByVal lIndex As Long = -1, _
       Optional ByVal vItemData As Variant, Optional vTag As Variant _
      ) As String
'''************************************************************************
'''Called By:     M_SynchList
'''
'''Arguments:
''' sItem         Text representation of the Item
''' lIndex
'''   Optional. Integer specifying the position within the object
'''   where the new item or row is placed
'''   If you supply a valid value(>= 0) for index, item is placed at that position within the object.
'''   If index is omitted, item is added at the proper sorted position
'''   (if the Sorted property is set to True) or to the end of the list (if Sorted is set to False).
''' vItemData
'''   Optional Associated variant data value
''' vTag
'''   Optional Additional value or object associated with the Item
'''
'''Description:   Adds List Item
'''
'''Note:          pctlData must be Listbox or ComboBox
'''               Unlike ItemData of the standard VB list controls
'''               vItemData is a variant and can Accept any data type
'''
'''Warning:
'''Updates Globals:
'''
'''Returns:       If successful  - Key value of the item
'''               Else           - Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_ListAddItem"

On Error GoTo ErrHnd
   
10    If mobjList Is Nothing Then
20       Set mobjList = New mclsLibList
      End If
      '''Generate Key value
80    mlListKey = mlListKey + 1
      '''Store new Item
100   Call mobjList.P_sAddItem(mlListKey, sItem, vItemData)
200   If VarType(vTag) = vbObject Then
220      Set mobjList.pobjItemNew.pvTag = vTag
      Else
300      mobjList.pobjItemNew.pvTag = vTag
      End If
      
400   If mlControlTypeData = lCT_List Then
         '''Assumed to be List or combo box
500      With mctlData
            '''Add Item to the list control
520         If lIndex >= 0 Then
540            Call .AddItem(sItem, lIndex)
560         Else
600            Call .AddItem(sItem)
620         End If
700         .ItemData(.NewIndex) = mlListKey
         End With
      End If
   
999   P_sListAddItem = CStr(mlListKey)
   
Exit Function
ErrHnd:
   P_sListAddItem = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Sub P_ListClear()
'''************************************************************************
'''Called By:     M_SynchList
'''               P_bLoadList
'''Arguments:
'''
'''Description:   Clears the list
'''
'''Note:          pctlData must be Listbox or ComboBox
'''
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_ListClear"

On Error GoTo ErrHnd
   
   If mlControlTypeData = lCT_List Then
      Call mctlData.Clear
   End If
   If Not mobjList Is Nothing Then
      Call mobjList.P_Clear
   Else
      Set mobjList = New mclsLibList
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Sub

Public Sub P_ListRemoveItem(Optional ByVal lIndex As Long = -1, Optional ByVal vItemData As Variant)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lIndex
'''   Index of the Item requested to be removed. Ignored if less then zero
''' vItemData
'''   Item data value of the Item requested to be removed
'''
'''Description:   Removes specified List Item
'''
'''Note:
'''
'''Warning:    if vItemData is specified all matching items will be removed
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_ListRemoveItem"
   
   Dim objItem          As mclsLibItem
   Dim sKey             As String
   
On Error GoTo ErrHnd
   
      '''If Remove by index is requested
      If lIndex >= 0 Then
100      sKey = mctlData.ItemData(lIndex)                   '''Get key value of the Item
120      Call mctlData.Remove(lIndex)                       '''Remove Item from list control
140      Call mobjList.P_RemoveItem(sKey)                   '''Remove item from collection
      End If
      '''If requested to remove all items matching specified Item data
200   If Not IsMissing(vItemData) Then
         Do
220         If mlControlTypeData = lCT_List Then
240            lIndex = P_lIndexByItemData(vItemData, sKey) '''Get index and key values of the Item
               If lIndex >= 0 Then
300               sKey = mctlData.ItemData(lIndex)
320               Call mctlData.RemoveItem(lIndex)              '''Remove Item from list control
               End If
            Else
400            sKey = mobjList.P_sFindByValue(vItemData)    '''Get key value of the Item
            End If
500         If Len(sKey) > 0 Then
520            Call mobjList.P_RemoveItem(sKey)             '''Remove item from collection
540            sKey = ""
            End If
         Loop Until Len(sKey) = 0
      End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Sub

Public Property Let pbDescAsToolTip(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE  Indicates whether to use Item Description for tool tip
'''
'''Note:
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pbDescAsToolTip"

On Error GoTo ErrHnd

   mbDescAsToolTip = bVal

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   
End Property

Public Property Get pbDescAsToolTip() As Boolean
   pbDescAsToolTip = mbDescAsToolTip
End Property

Public Property Get pbEditable() As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE  Indicates whether end user can modify the value
'''
'''
'''Note:          Does not effect pobjDataField.pbEditable - value can still be changed by
'''               the code
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - current value
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pbEditable"

On Error GoTo ErrHnd

   pbEditable = mobjItem.pbEditable

Exit Property
ErrHnd:
   pbEditable = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Let pbEditable(ByVal bValue As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
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
   Const sPROC_NAME     As String = "Let pbEditable"

On Error GoTo ErrHnd

   mobjItem.pbEditable = bValue
   Call M_ApplyControlProps
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get pbListItemValueExist(ByVal vVal As Variant) As Boolean
   pbListItemValueExist = mobjList.pbExistValue(ByVal vVal)
   
End Property

Public Property Get pbVisible() As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
'''               Indicates whether associated GUI controls are visible
'''
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - current value
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pbVisible"

On Error GoTo ErrHnd

   pbVisible = mbVisible

Exit Property
ErrHnd:
   pbVisible = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Let pbVisible(ByVal bValue As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
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
   Const sPROC_NAME     As String = "Let pbVisible"

On Error GoTo ErrHnd

   mbVisible = bValue

   If mbControlCaptionSet Then
      mctlCaption.Visible = mbVisible
   End If
   
   If Not mctlData Is Nothing Then
      mctlData.Visible = mbVisible
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get pctlCaption() As Object
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
'''               Y  Y  Nothing
'''               Control associated with caption
'''
'''Note:          Can be same as one associated with ctlData
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - current value
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pctlCaption"

On Error GoTo ErrHnd

   Set pctlCaption = mctlCaption

Exit Property
ErrHnd:
   pctlCaption = Nothing
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Set pctlCaption(ByVal ctlCaption As Object)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  Nothing
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
   Const sPROC_NAME     As String = "Set pctlCaption"

On Error GoTo ErrHnd

   Set mctlCaption = ctlCaption
   mbControlCaptionSet = Not mctlCaption Is Nothing
   If mbControlCaptionSet Then
      If psCaption = "" Then
         psCaption = mctlCaption.Caption
      End If
   End If
Exit Property
ErrHnd:
   mbControlCaptionSet = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get pctlData() As Object
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  Nothing
'''               Control associated with value
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - current value
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pctlData"

On Error GoTo ErrHnd

   Set pctlData = mctlData

Exit Property
ErrHnd:
   pctlData = Nothing
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Set pctlData(ctlData As Object)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  Nothing
'''
'''Note:
'''
'''Warning:       If specified ctlData contains reference to List control - this procedure
'''               will attemt to preserve contents of the specified list(note: Item and Item data only)
'''               as well as existing list.
'''               The result will be combine list from both sources.
'''               Therefore clear list(s) that you do not want to preserve prior to calling this procedure.
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Set pctlData"

On Error GoTo ErrHnd

   mlControlTypeData = lCT_Undefined
   Set mctlData = ctlData
   
   If Not mctlData Is Nothing Then
      Select Case True
      Case TypeOf mctlData Is CheckBox
         mlControlTypeData = lCT_Check
      Case TypeOf mctlData Is OptionButton
         mlControlTypeData = lCT_Option
      Case TypeOf mctlData Is TextBox
         mlControlTypeData = lCT_Text
      Case TypeOf mctlData Is Label
         mlControlTypeData = lCT_Label
      Case (TypeOf mctlData Is ComboBox) Or (TypeOf mctlData Is ListBox)
         mlControlTypeData = lCT_List
         Call M_SynchList  '''Make sure that associated collection is in synch
      Case (TypeOf mctlData Is DTPicker)
         mlControlTypeData = lCT_DTPicker
      Case Else
         mlControlTypeData = lCT_Undefined
      End Select
      Call M_ApplyControlProps   '''Make sure that property settings are reflected in the GUI
   End If
   
Exit Property
ErrHnd:
   mlControlTypeData = lCT_Undefined
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get plListCount() As Long
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:   Return current List count
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if successful
'''                  Current list count
'''               else
'''                  -1
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plListCount"

   Dim lResult          As Long
   
On Error GoTo ErrHnd

20    If Not mobjList Is Nothing Then
40       lResult = mobjList.plCount
      End If
      
999   plListCount = lResult
      
Exit Property
ErrHnd:
   plListCount = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Property Get pobjItem() As mclsLibItem
   Set pobjItem = mobjItem
End Property

Public Property Get pobjListItem(ByVal lIdx As Long) As mclsLibItem
   
   Set pobjListItem = mobjList.pobjItem(lIdx)

End Property

Public Property Get pobjListItemByValue(ByVal vItemValue As Variant) As mclsLibItem
   Set pobjListItemByValue = mobjListItem(vItemValue)
End Property

Public Property Get pobjListItemCurrent() As mclsLibItem
   Set pobjListItemCurrent = mobjListItem
End Property

Public Property Get psFormat() As String
   psFormat = msFormat
End Property

Public Property Let psFormat(ByVal sVal As String)
   msFormat = sVal
   Call P_bRefresh
End Property


Public Property Get psListItem _
      (Optional ByRef vTag As Variant, Optional ByVal vByItemData As Variant) As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' vTag
'''   Optional tag associated with the currently selected item requested to be returned byref
''' vByItemData
'''   Optional Identifier of requested Item.
'''   If not specified properties of the current item will be returned
'''
'''Description:   Return Name of the requested list item and
'''               Tag associated with it
'''
'''Note:          If List control is set the properties of the current item may not match
'''               current pobjItem.pvValue unless P_AcceptControlValue is called
'''               Also refer to Get pvValue.
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if successful
'''                  Requested list item
'''               else
'''                  Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psListItem"

   Dim sResult          As String
   Dim objItem          As mclsLibItem
   
On Error GoTo ErrHnd
      
100   Set objItem = mobjListItem(vByItemData)
      
400   If Not objItem Is Nothing Then
420      sResult = objItem.psName
440      If VarType(objItem.pvTag) = vbObject Then
460         Set vTag = objItem.pvTag
         Else
480         vTag = objItem.pvTag
500         If IsMissing(vTag) Then '''VB flick
520            vTag = Empty
            End If
         End If
      End If

999   psListItem = sResult
      
Exit Property
ErrHnd:
   psListItem = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Property Get psValueFormated() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:   Return curent value in current format
'''
'''
'''Note:
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if successful
'''                  curent value in current format
'''               else
'''                  Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psValueFormated"
  
On Error GoTo ErrHnd

999   psValueFormated = M_sFormatedValue(mobjItem.pvValue)
   
Exit Property
ErrHnd:
   psValueFormated = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Property Get pvTag(Optional ByVal vByItemData As Variant) As Variant
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' vByItemData
'''   Optional Identifier of requested Item.
'''   If not specified tag of the current item will be returned
'''
'''Description:   Return tag of the requested list item
'''
'''
'''Note:          If List control is set the properties of the current item may not match
'''               current pobjItem.pvValue unless P_AcceptControlValue is called
'''               Also refer to Get pvValue.
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if successful
'''                  Requested list item tag
'''               else
'''                  Empty
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pvTag"

   Dim vResult          As Variant
   
On Error GoTo ErrHnd
      
20    If Len(psListItem(vResult, vByItemData)) = 0 Then
40       vResult = Empty
      Else
100      If VarType(vResult) = vbObject Then
            Set pvTag = vResult
         Else
200         pvTag = vResult
         End If
      End If

Exit Property
ErrHnd:
   pvTag = Empty
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Property Let pvValue(vValue As Variant)
'''************************************************************************
'''Called By:     P_bRefresh or any
'''
'''Arguments:
''' vValue        New value
'''
'''
'''Description:   Update field value based on the current data control type
'''
'''Note:          This procedure will only change display value - it will not change item or data value
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pvValue"

   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd

   If Not bIn_static Then
      bIn_static = True
400   Call M_bDisplayValue(vValue)
500   If pbAutoAcceptControlValue Then
520      Call P_AcceptControlValue
      End If

      bIn_static = False
   End If
   
Exit Property
ErrHnd:
   bIn_static = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, psCaption & vbCrLf & gsLINE_ & Erl, True)

End Property

Public Property Get pvValue() As Variant
'''************************************************************************
'''Called By:     P_bRefresh or any
'''
'''Arguments:
'''
'''
'''Description:   Return current field value
'''
'''Note:
'''
'''Warning:       Will cause an ERROR if Data Control is not set - use pobjItem.pvValue or psValueFormated instead
'''
''' GUI Field value will returns current control value which may not match
''' Item value(e.g. pbAutoAcceptControlValue is false).
''' Use P_AcceptControlValue or assign value directly as
''' objLibGUIField.pobjItem.pvValue = objLibGUIField.pvValue to accept control value.
''' This is done to allow to validate the value entered by the user prior
''' to overiding current value.
''' Usage:
'''                  if objLibGUIField.pobjItem.pvValue is valid then
'''                     objLibGUIField.P_AcceptControlValue
'''                  else
'''                     objLibGUIField.P_bRefresh
'''Updates Globals:
'''
'''Returns:       if successful
'''                  current value
'''               else
'''                  Empty
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pvValue"

   Static bIn_static    As Boolean
   Dim vVal             As Variant
   
On Error GoTo ErrHnd
      
220   If (Not bIn_static) Then
         Select Case mlControlTypeData
            Case lCT_Check
100            vVal = mctlData.Value = 1
            Case lCT_Option
200            vVal = mctlData.Value
            Case lCT_Text
300            vVal = gobjLibValid.LibValid_vCastType(mctlData.Text)
            Case lCT_Label
400            vVal = gobjLibValid.LibValid_vCastType(mctlData.Caption)
            Case lCT_List
500            vVal = M_vGetCurrentItemData()
            Case lCT_DTPicker
600            vVal = mctlData.Value
            Case Else
               vVal = mobjItem.pvValue
'''<DSN 1000           Call gobjLibApp.LibApp_Error _
               (0, gobjLibStr.LibStr_sBuildSentence _
                (gsCONTROL, gsTYPE, TypeName(mctlData), gsIS, gsNOT, gsSUPPORTED), _
                msMODULE_NAME, sPROC_NAME, "", True, False, False _
               )
         End Select

1000     If pbAutoAcceptControlValue Then
1020        Call P_AcceptControlValue
         End If
         
         bIn_static = False
      End If

9999  pvValue = vVal
   
Exit Property
ErrHnd:
   pvValue = Empty
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)
   bIn_static = False
   
End Property

Private Sub Class_Initialize()
'''************************************************************************
'''Called By:     None
'''
'''Arguments
'''
'''Description:   Initialize object
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
   Const sPROC_NAME     As String = "Class_Initialize"

On Error GoTo ErrHnd

   '''Setup default values
20 mbVisible = True
30 Set mobjItem = New mclsLibItem
40 mbDescAsToolTip = True
60 mbEnabled = True
80 pbAutoUpdate = True

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Sub


Private Sub Class_Terminate()
'''************************************************************************
'''Called By:     None
'''
'''Arguments
'''
'''Description:   Release object
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
   Const sPROC_NAME     As String = "Class_Terminate"

On Error GoTo ErrHnd

   Set mobjItem = Nothing
   Set mctlCaption = Nothing
   Set mctlData = Nothing
   Set mobjList = Nothing
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Sub
Public Property Get pbEnabled() As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
'''               Indicates whether note can be modified - grayed out
'''               Propagated - will apply to all child notes.
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       True  - if All child notes are enabled
'''               False - else or error
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pbEnabled"

On Error GoTo ErrHnd

   pbEnabled = mbEnabled
   
Exit Property
ErrHnd:
   pbEnabled = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Let pbEnabled(ByVal bValue As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
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
   Const sPROC_NAME     As String = "Let pbEnabled"

On Error GoTo ErrHnd

   mbEnabled = bValue
   Call M_ApplyControlProps
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get psCaption() As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  Empty
'''               Presentation Name of the field
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - current value
'''               Else           - empty
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psCaption"

On Error GoTo ErrHnd

   psCaption = mobjItem.psName

Exit Property
ErrHnd:
   psCaption = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Let psCaption(ByVal sValue As String)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W  Default
'''               Y  Y  TRUE
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
   Const sPROC_NAME     As String = "Let psCaption"

   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd

   If Not bIn_static Then
      bIn_static = True
      mobjItem.psName = sValue
      Call M_bDisplayCaption
      bIn_static = False
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   bIn_static = False
   
End Property

Private Sub mobjItem_pevBeforeGetValue()
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''
'''Description:   If required update item value before it returns it
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
   Const sPROC_NAME     As String = "mobjItem_pevBeforeGetValue"

On Error GoTo ErrHnd

500   If pbAutoAcceptControlValue Then
520      Call P_AcceptControlValue
      End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, psCaption & vbCrLf & gsLINE_ & Erl)

End Sub

Private Sub mobjItem_pevEditableChanged(ByVal bNewVal As Boolean)
   Call M_ApplyControlProps '''Make sure that property settings are reflected in the GUI
End Sub

Private Sub mobjItem_pevNameChanged()
   If pbAutoUpdate Then
      Call P_bRefresh
   End If
End Sub

Private Sub mobjItem_pevValueChanged()
   If pbAutoUpdate Then
      Call P_bRefresh
   End If
End Sub


