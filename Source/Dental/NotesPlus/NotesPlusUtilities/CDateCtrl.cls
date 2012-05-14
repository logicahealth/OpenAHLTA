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
Attribute VB_Name = "CDateCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrDate As String
Private mstrFormCaption As String
Private mblnAllowFutureDate As Boolean

Public Enum TimePeriodEnums
    PastYear = 0
    Recently = 1
    Childhood = 2
End Enum
Public Property Get UserDateTime() As String
    UserDateTime = mstrDate
End Property
Public Property Let FormCaption(ByVal strCaption As String)
    mstrFormCaption = strCaption
End Property
Public Property Get FormCaption() As String
    FormCaption = mstrFormCaption
End Property

Public Property Let AllowFutureDate(ByVal blnVal As Boolean)
    mblnAllowFutureDate = blnVal
End Property
Public Property Get AllowFutureDate() As Boolean
    AllowFutureDate = mblnAllowFutureDate
End Property

Public Function Get_User_Input(ByVal strDate As String) As Boolean
    Dim blnReturn As Boolean
    
    If strDate <> "" Then
        frmDateSelector.Set_Date strDate
    End If
        
    If mstrFormCaption <> "" Then
        frmDateSelector.Caption = mstrFormCaption
    End If
    
    'KDunne 9/4/2003
    'SCR 39717 - Sometimes we do not want to allow a future date to be selected.
    '
    frmDateSelector.AllowFutureDate = mblnAllowFutureDate
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmDateSelector.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    
    If frmDateSelector.UserCancelled Then
        blnReturn = False
    Else
        blnReturn = True
        mstrDate = frmDateSelector.UserDateTime
        
        Unload frmDateSelector
    End If
    
    Get_User_Input = blnReturn
End Function

Private Sub Class_Initialize()
    'KDunne 9/4/2003
    'SCR 39717
    mblnAllowFutureDate = True
End Sub
