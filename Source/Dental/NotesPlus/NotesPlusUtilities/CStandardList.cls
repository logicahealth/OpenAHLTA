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
Attribute VB_Name = "CStandardList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDentalMLParent

Private mstrFormCaption As String
Private mobjItem As CMultiListItem
Private mobjListItems As CMultiListItems
Private mColItems As Collection
Private mblnCancelled As Boolean
Private mstrValues As String

Private Const const_MODULE_NAME = "CStandardList"


Public Function ValueStringDisplay() As String
    
    ValueStringDisplay = mobjListItems.Value_String_Display
    
End Function


Private Sub Class_Initialize()
    IntializeMultiListItems
    mblnCancelled = False
End Sub

Private Sub Class_Terminate()
    Set mobjListItems = Nothing
    Set mobjItem = Nothing
End Sub

Public Property Let Form_Caption(ByVal strCaption As String)
    mstrFormCaption = strCaption
End Property

Public Property Get Form_Caption() As String
    Form_Caption = mstrFormCaption
End Property

Public Property Get MultiListItems() As CMultiListItems
    Set MultiListItems = mobjListItems
End Property

Public Property Set MultiListItems(ByVal objMultiList As CMultiListItems)
    Set mobjListItems = objMultiList
End Property

Public Property Get ValueString() As String
    ValueString = mstrValues
End Property

'*******************************************************************
'Function: Load_String
'Author: KDunne
'Date: 10/08/2001
'Returns: boolean. True if successful, False if error occurs
'Purpose: Populates internal collection from | deliminated string
'*******************************************************************
Public Function Load_String(ByVal strValues As String) As Boolean
    
    IntializeMultiListItems
    
    Set mobjItem = New CMultiListItem
    
    If strValues = "" Then
        Load_String = False
    Else
        mstrValues = strValues
        Load_String = mobjListItems.LoadFromValueString(strValues)
    End If
    
End Function

'*******************************************************************
'Function: Get_User_Input
'Author: KDunne
'Date: 10/08/2001
'Returns: boolean. True if successful, False if error occurs or
'         User cancels.
'Purpose: Prompts user for Input. Displays checkboxes, descriptive text.
'         After user selects "OK", a | deliminated value string is generated.
'*******************************************************************
Public Function Get_User_Input() As Boolean
    Dim blnReturn As Boolean
    
    Set frmStandard.Parent = Me
    Set frmStandard.MultiListItems = mobjListItems
    
    mblnCancelled = False
    frmStandard.Caption = mstrFormCaption

    blnReturn = frmStandard.Build_Form_Controls
    If blnReturn Then
        Load frmStandard
        cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        frmStandard.Show vbModal
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        
        If mblnCancelled Then
            blnReturn = False
        Else
            blnReturn = True
            mstrValues = mobjListItems.BuildValueString
        End If
    End If
    
    Get_User_Input = blnReturn
   
End Function

'*******************************************************************
'Sub: Add
'Author: KDunne
'Date: 10/8/2001
'Purpose: Adds a CMultiListItem to the CMultiListItems collection class.
'*******************************************************************
Public Sub Add(ByVal strKey As String, ByVal strDescription As String, ByVal blnChecked As Boolean)
            
    If mobjListItems Is Nothing Then
        IntializeMultiListItems
    End If
    
    If mobjListItems.Count >= const_MAX_LIST_ITEMS Then
        Err.Raise const_ERR_MAX_ELEMENTS, const_MODULE_NAME & ":Add()", "Maximum number of elements is " & const_MAX_LIST_ITEMS
    Else
        mobjListItems.Add strKey, strDescription, blnChecked
    End If
End Sub

'*******************************************************************
'Sub: IntializeMenuListItems
'Author: Kdunne
'Date: 10/8/2001
'Purpose: Creates a new instance of CMultiListItems and sets the
'         objects ListType property to Standard.
'*******************************************************************
Private Sub IntializeMultiListItems()
    Set mobjListItems = New CMultiListItems
    mobjListItems.ListType = StandardList
End Sub

Private Sub IDentalMLParent_UserCancelled()
    mblnCancelled = True
End Sub
