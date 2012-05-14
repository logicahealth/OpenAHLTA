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
Attribute VB_Name = "Mnu_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarButtons As colButtons
Private mvarMenuItems As colMenuItems
Private mvarSubMenuItems As colSubMenuItems
Private iMenuBtnID As Integer
Private mvarColRemovedLocations As Collection

'SF Not used in this class
Public currentObjMnu As clsButton
Public currentObjMnuItem As clsMenuItem
'

Public Enum eGroup
  eDx = 0
  eTx = 1
  eTxP = 2
  eDefTx = 3
  eWDx = 4
End Enum

Public Enum eToothMode
  eFullMouth = 0
  eBigTooth = 1
End Enum

Public Enum eViewMode
  eDxView = 0
  eTxView = 1
  eDx603aView = 2
End Enum

Public Property Set MenuItems(ByVal vData As colMenuItems)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MenuItems = Form1
    Set mvarMenuItems = vData
End Property


Public Property Get MenuItems() As colMenuItems
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuItems
    Set MenuItems = mvarMenuItems
End Property
Public Property Set RemovedLocations(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MenuItems = Form1
    Set mvarColRemovedLocations = vData
End Property


Public Property Get RemovedLocations() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuItems
    Set RemovedLocations = mvarColRemovedLocations
End Property
Public Property Set SubMenuItems(ByVal vData As colSubMenuItems)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MenuItems = Form1
    Set mvarSubMenuItems = vData
End Property


Public Property Get SubMenuItems() As colSubMenuItems
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuItems
    Set SubMenuItems = mvarSubMenuItems
End Property


Public Property Set Buttons(ByVal vData As colButtons)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Menus = Form1
    Set mvarButtons = vData
End Property


Public Property Get Buttons() As colButtons
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Menus
    Set Buttons = mvarButtons
End Property
Private Sub loadMenuItemDef()
  Dim strUniMenuItems As String
  Dim strMenuItems As String
  Dim strArrMenuItemRows() As String
  Dim strArrMenuItemColumns() As String
 
  Dim objMnu As clsButton
  Dim objMnuItemInstr As clsMenuItemInstr
  Dim objMnuItem As clsMenuItem
  
  Dim i As Integer
  Dim eBtnGroup As eGroup
  
   strUniMenuItems = LoadResData("MENUITEMS", "CUSTOM")
   strMenuItems = StrConv(strUniMenuItems, vbUnicode)
   strArrMenuItemRows = Split(strMenuItems, vbCrLf)
   
   For i = 0 To UBound(strArrMenuItemRows)

        strArrMenuItemColumns = Split(strArrMenuItemRows(i), "|")

            Set objMnuItem = New clsMenuItem
            
            If UBound(strArrMenuItemColumns) = 0 Or UBound(strArrMenuItemColumns) = -1 Then Exit For 'No More Columns
            
            
            objMnuItem.MnuButtonID = strArrMenuItemColumns(1)
            objMnuItem.MnuButtonName = strArrMenuItemColumns(3)
            objMnuItem.MnuSection = strArrMenuItemColumns(2)
            objMnuItem.MnuGroup = strArrMenuItemColumns(0)

            objMnuItem.mnuItemId = strArrMenuItemColumns(4)
            objMnuItem.MnuItemName = Trim(strArrMenuItemColumns(5))
            objMnuItem.MnuItemCodeID = strArrMenuItemColumns(6)
            objMnuItem.MnuItemToInstructionID = strArrMenuItemColumns(17)
            objMnuItem.SubMnuItemID = strArrMenuItemColumns(18)
            objMnuItem.MnuItemVisible = True
            
            'Set the biz rules
            objMnuItem.BizRule = strArrMenuItemColumns(7)
            objMnuItem.BizSurface = strArrMenuItemColumns(8)
            objMnuItem.BizSurfacePlus = strArrMenuItemColumns(9)
            objMnuItem.BizToothSequence = strArrMenuItemColumns(10)
            objMnuItem.BizToothPosition = strArrMenuItemColumns(11)
            objMnuItem.BizToothType = strArrMenuItemColumns(12)
            objMnuItem.BizToothEndo = strArrMenuItemColumns(13)
            objMnuItem.BizToothMax = strArrMenuItemColumns(14)
            objMnuItem.BizImplantOrTooth = strArrMenuItemColumns(15)
            objMnuItem.BizToothRoots = strArrMenuItemColumns(16)
            'SCR 36412
            objMnuItem.BizToothSupernumerary = strArrMenuItemColumns(19)
            
            If objMnuItem.MnuButtonID <> iMenuBtnID Then  'New Button found
            
                iMenuBtnID = objMnuItem.MnuButtonID
                Set objMnu = New clsButton
                
                objMnu.MnuButtonID = objMnuItem.MnuButtonID
                objMnu.MnuButtonName = objMnuItem.MnuButtonName
                objMnu.MnuSection = objMnuItem.MnuSection
                objMnu.MnuGroup = objMnuItem.MnuGroup
                
                mvarButtons.Add objMnu, CStr(objMnu.MnuButtonID)
                
                mvarButtons(CStr(objMnu.MnuButtonID)).MnuItems.Add objMnuItem, CStr(objMnuItem.mnuItemId)
            Else
                mvarButtons(CStr(objMnu.MnuButtonID)).MnuItems.Add objMnuItem, CStr(objMnuItem.mnuItemId)
            End If
        
            mvarMenuItems.Add objMnuItem, CStr(objMnuItem.mnuItemId)

   Next i
   
End Sub
Private Sub loadSubMenuItemDef()
  Dim strUniSubMenuItems As String
  Dim strSubMenuItems As String
  Dim strArrSubMenuItemRows() As String
  Dim strArrSubMenuItemColumns() As String
 
   Dim objSubMnuItem As clsSubMenuItem
  
  Dim i As Integer
  Dim eBtnGroup As eGroup
  
   strUniSubMenuItems = LoadResData("SUBMENUITEMS", "CUSTOM")
   strSubMenuItems = StrConv(strUniSubMenuItems, vbUnicode)
   strArrSubMenuItemRows = Split(strSubMenuItems, vbCrLf)
   
   For i = 0 To UBound(strArrSubMenuItemRows)

        strArrSubMenuItemColumns = Split(strArrSubMenuItemRows(i), "|")

            Set objSubMnuItem = New clsSubMenuItem
            
            If UBound(strArrSubMenuItemColumns) = 0 Or UBound(strArrSubMenuItemColumns) = -1 Then Exit For 'No More Columns
            
            
            objSubMnuItem.SubMnuID = strArrSubMenuItemColumns(0)
            objSubMnuItem.mnuItemId = strArrSubMenuItemColumns(1)
            objSubMnuItem.SubMnuItemID = strArrSubMenuItemColumns(2)
     
            
            mvarSubMenuItems.Add objSubMnuItem, CStr(objSubMnuItem.SubMnuID)

   Next i
   
End Sub
Private Sub loadRemovedLocations()
  Dim strUniRemovedLocItems As String
  Dim strRemovedLocItems As String
  Dim strArrRemovedLocItemRows() As String
  Dim strArrRemovedLocItemColumns() As String
 
   Dim objRemovedLocItem As clsRemovedLocations
  
  Dim i As Integer
  
  On Error Resume Next
  
   strUniRemovedLocItems = LoadResData("REMOVEDLOCATIONS", "CUSTOM")
   strRemovedLocItems = StrConv(strUniRemovedLocItems, vbUnicode)
   strArrRemovedLocItemRows = Split(strRemovedLocItems, vbCrLf)
   
   For i = 0 To UBound(strArrRemovedLocItemRows)

        strArrRemovedLocItemColumns = Split(strArrRemovedLocItemRows(i), ",")

            Set objRemovedLocItem = New clsRemovedLocations
            
            If UBound(strArrRemovedLocItemColumns) = 0 Or UBound(strArrRemovedLocItemColumns) = -1 Then Exit For 'No More Columns
            
            
            objRemovedLocItem.mnuItemId = strArrRemovedLocItemColumns(0)
            objRemovedLocItem.eRemovedLocation = strArrRemovedLocItemColumns(1)
           
            mvarColRemovedLocations.Add objRemovedLocItem, CStr(objRemovedLocItem.mnuItemId)

   Next i
   
End Sub
Private Sub LoadMenuInstructions()

    'This sub reads the resource file and loads the menuinstruction collection
Dim strUniMenuInstrItems As String
Dim strMenuInstrItems As String
Dim strArrMenuInstrItemRows() As String
Dim strArrMenuInstrItemColumns() As String
 
Dim oMenuInstruction As clsMenuItemInstr
  
Dim i As Integer

On Error GoTo ErrHnd

   strUniMenuInstrItems = LoadResData("MENUINSTRUCTIONS", "CUSTOM")
   strMenuInstrItems = StrConv(strUniMenuInstrItems, vbUnicode)
   strArrMenuInstrItemRows = Split(strMenuInstrItems, vbCrLf)
   
   Set goMenuInstructions = New colMenuItemInstrs
   
   For i = 0 To UBound(strArrMenuInstrItemRows)

        strArrMenuInstrItemColumns = Split(strArrMenuInstrItemRows(i), ",")

            
            If UBound(strArrMenuInstrItemColumns) = 0 Or UBound(strArrMenuInstrItemColumns) = -1 Then Exit For 'No More Columns
         
            Set oMenuInstruction = New clsMenuItemInstr

         
        
        
        With oMenuInstruction
            .CodeID = strArrMenuInstrItemColumns(0)
            .Group = strArrMenuInstrItemColumns(1)
            .MenuInstructionID = strArrMenuInstrItemColumns(2)
            .Color = strArrMenuInstrItemColumns(3)
            .Hatch = strArrMenuInstrItemColumns(4)
            .Symbol = strArrMenuInstrItemColumns(5)
            .Operation = strArrMenuInstrItemColumns(6)
            .Location = strArrMenuInstrItemColumns(7)
            .RemovedLocations = strArrMenuInstrItemColumns(8)
        End With
        
        'Fill the collection with the correct key
        goMenuInstructions.Add oMenuInstruction, oMenuInstruction.CodeID & "|" & oMenuInstruction.Group
    Next

Exit Sub
ErrHnd:
    MsgBxARMd Err.Description & " source:Mnu_Lib.LoadMenuInstruction. Could not load the menu instruction in the collection.", vbCritical
        
End Sub
Private Sub Class_Initialize()
    Set mvarButtons = New colButtons
    Set mvarMenuItems = New colMenuItems
    Set mvarSubMenuItems = New colSubMenuItems
    Set mvarColRemovedLocations = New Collection
    
    loadMenuItemDef
    loadSubMenuItemDef
    'Init and load the menu instructions
    Set goMenuInstructions = New colMenuItemInstrs
    
    LoadMenuInstructions
    loadRemovedLocations
    
       '--- SCR 28371 RWalker 10/10/02 Instantiate ARM / Init transactions
    Call ARM_InitAllTransactionsForDental_Dntl_Mnu_Lib

End Sub

Private Sub Class_Terminate()
    Set mvarButtons = Nothing
    Set mvarMenuItems = Nothing
    Set currentObjMnuItem = Nothing
    Set currentObjMnu = Nothing
End Sub
