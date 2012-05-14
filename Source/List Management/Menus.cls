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
Attribute VB_Name = "Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Manages Creation and Destruction of Menus."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"OneProblem"
Attribute VB_Ext_KEY = "Member0" ,"OneProblem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum EnumMenus
    MENU_ADD = &HFFFFFF
    MENU_SAVE
    MENU_CANCEL
    MENU_DELETE
    MENU_TEST
    MENU_EDIT
    MENU_REFRESH
    MENU_GUIDELINES
    MENU_LAYOUT
    MENU_SAVEAS
    MENU_NEW
    MENU_IMPORT
    MENU_EXPORT
    MENU_PRINT
    MENU_PRINTPREVIEW
    MENU_COPY_TO
    MENU_MOVE_DOWN
    MENU_MOVE_UP
End Enum

'local variable to hold collection
Private mCol As Collection
Private objCWComm As Comm

Public Sub Init(ByRef objComm As Comm)
    Set objCWComm = objComm
End Sub
'Set Menu State
Public Function SetMenuEnabled(ByVal Key As EnumMenus, ByVal Enabled As Boolean)

    Dim bRet As Boolean
    Const MENU_ENABLE As String = "ENABLE"
    Const MENU_DISABLE As String = "DISABLE"
    
    'jrm:Lock down designer
    If gbDesignerIsLocked = True Then
        If Not Key = MENU_REFRESH And Not Key = MENU_PRINT Then Exit Function
    End If
    
    If Enabled Then
        bRet = objCWComm.UpdateMenuItem(mCol(CStr(Key)), MENU_ENABLE)
    Else
        bRet = objCWComm.UpdateMenuItem(mCol(CStr(Key)), MENU_DISABLE)
    End If
    
End Function
'Create The Menus
Public Sub Add(ByVal Key As EnumMenus)
Attribute Add.VB_Description = "Use to add a new menu."
    Dim l As Long
    
    'jrm:Lock down designer
    If gbDesignerIsLocked = True Then
        If Not Key = MENU_REFRESH And Not Key = MENU_PRINT Then
            'Must add to the collection even though we will never call this since
            'other items down stream expect to have this populated
            mCol.Add l, CStr(Key)
            Exit Sub
        End If
    End If
    
    Select Case Key
    Case MENU_ADD
        l = objCWComm.AddActionMenuItem("&Add", "new.bmp", cwlMyProgId, "Add Item")
    Case MENU_NEW
        l = objCWComm.AddActionMenuItem("&New...", "new.bmp", cwlMyProgId, "New Item")
    Case MENU_EDIT
        l = objCWComm.AddActionMenuItem("&Edit...", "edit.bmp", cwlMyProgId, "Edit Item")
    Case MENU_SAVE
        l = objCWComm.AddActionMenuItem("&Save", "save.bmp", cwlMyProgId, "Save")
    Case MENU_COPY_TO
        l = objCWComm.AddActionMenuItem("&Copy To...", "saveas.bmp", cwlMyProgId, "Copy To...")
    Case MENU_SAVEAS
        l = objCWComm.AddActionMenuItem("&Save As...", "saveas.bmp", cwlMyProgId, "Save As...")
    Case MENU_CANCEL
        l = objCWComm.AddActionMenuItem("&Cancel", "cancel.bmp", cwlMyProgId, "Cancel")
    
    Case MENU_DELETE
        l = objCWComm.AddActionMenuItem("&Delete", "delete.bmp", cwlMyProgId, "Delete all problems menu item")
    Case MENU_TEST
        l = objCWComm.AddActionMenuItem("&Test", "cancel.bmp", cwlMyProgId, "Test menu item")
    Case MENU_REFRESH
        l = objCWComm.AddActionMenuItem("&Refresh", "refresh.bmp", cwlMyProgId, "Refresh Display")
    Case MENU_GUIDELINES
        l = objCWComm.AddActionMenuItem("&Guidelines", "wellness.bmp", cwlMyProgId, "Show Wellness Reminder Guidelines")
    Case MENU_LAYOUT
        l = objCWComm.AddActionMenuItem("&Layout", "copy_template.bmp", cwlMyProgId, "")
    Case MENU_IMPORT
        l = objCWComm.AddActionMenuItem("&Import...", "import_template.bmp", cwlMyProgId, "Import...")
    Case MENU_EXPORT
        l = objCWComm.AddActionMenuItem("&Export...", "export_template.bmp", cwlMyProgId, "Export...")
    Case MENU_PRINT
        l = objCWComm.AddPrintMenuItem("&Print...", cwlMyProgId, "Print...")
    Case MENU_PRINTPREVIEW
        l = objCWComm.AddPrintPreviewMenuItem("&Print Preview...", cwlMyProgId, "Print...")
    Case MENU_MOVE_DOWN
        l = objCWComm.AddActionMenuItem("&Move Down...", "down.bmp", cwlMyProgId, "Move item down in list...")
    Case MENU_MOVE_UP
        l = objCWComm.AddActionMenuItem("&Move Up...", "up.bmp", cwlMyProgId, "Move item up in list...")
    End Select
    
    
    mCol.Add l, CStr(Key)
Debug.Print "menu add:" & Key
End Sub
Public Sub Remove(vntIndexKey As EnumMenus)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    objCWComm.DeleteMenuItem CStr(vntIndexKey)

End Sub
Public Sub RemoveAll()
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    Dim l
    
    For Each l In mCol
        objCWComm.DeleteMenuItem l
    Next

End Sub

Public Property Get Item(vntIndexKey As EnumMenus) As Long
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
    Item = mCol(CStr(vntIndexKey))
End Property


Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    'destroys all menus not explicitly destroyed
    Dim l As Variant
    
    For Each l In mCol
        objCWComm.DeleteMenuItem l
    Next

    Set mCol = Nothing
    
End Sub

