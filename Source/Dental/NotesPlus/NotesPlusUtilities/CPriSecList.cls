VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPriSecList"
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

Private Const const_MODULE_NAME = "CPriSecList"

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
Friend Sub User_Cancelled()
    mblnCancelled = True
End Sub


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
'Purpose: Prompts user for Input. Displays checkboxes, descriptive text
'         a primary radio button and a secondary checkbox for each item.
'         There can be only one primary item, but there can be many secondary
'         items. An item can only be a primary or secondary item, not both.
'*******************************************************************
Public Function Get_User_Input() As Boolean
    Dim blnReturn As Boolean
    
    Set frmPriSec.Parent = Me
    Set frmPriSec.MultiListItems = mobjListItems
    
    frmPriSec.Caption = mstrFormCaption
    
    mblnCancelled = False
    
    blnReturn = frmPriSec.Build_Form_Controls
    If blnReturn Then
        Load frmPriSec
        
        cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
        frmPriSec.Show vbModal
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
Public Sub Add(ByVal strKey As String, ByVal strDescription As String, _
    ByVal blnChecked As Boolean, ByVal blnPrimary As Boolean, _
    ByVal blnSecondary As Boolean)
            
    If mobjListItems Is Nothing Then
        IntializeMultiListItems
    End If
    
    If mobjListItems.Count >= const_MAX_LIST_ITEMS Then
        Err.Raise const_ERR_MAX_ELEMENTS, const_MODULE_NAME & ":Add()", "Maximum number of elements is " & const_MAX_LIST_ITEMS
    Else
        mobjListItems.Add strKey, strDescription, blnChecked, blnPrimary, blnSecondary
    End If
End Sub

'*******************************************************************
'Sub: IntializeMenuListItems
'Author: Kdunne
'Date: 10/8/2001
'Purpose: Creates a new instance of CMultiListItems and sets the
'         objects ListType property to PrimarySecondaryList.
'*******************************************************************
Private Sub IntializeMultiListItems()
    Set mobjListItems = New CMultiListItems
    mobjListItems.ListType = PrimarySecondaryList
End Sub

Private Sub IDentalMLParent_UserCancelled()
    mblnCancelled = True
End Sub
