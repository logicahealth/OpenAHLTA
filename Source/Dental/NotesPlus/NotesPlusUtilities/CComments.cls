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
Attribute VB_Name = "CComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IComments

Private mstrComments As String
Private mlngMaxLength As Long
Private mstrCaption As String
Private mblnCancelled As Boolean
Private mblnValidate As Boolean
Private mblnShowingModalForm As Boolean

Public Property Let Showing_Modal_Form(ByVal blnValue As Boolean)
    mblnShowingModalForm = blnValue
End Property
Public Property Get Showing_Modal_Form() As Boolean
    Showing_Modal_Form = mblnShowingModalForm
End Property

Public Property Let Validate_Text(ByVal blnValue As Boolean)
    mblnValidate = blnValue
End Property
Public Property Let Caption(ByVal strCaption As String)
    mstrCaption = strCaption
End Property
Public Property Get Caption() As String
    Caption = mstrCaption
End Property
Public Property Let Comments(ByVal strComments As String)
    mstrComments = strComments
End Property
Public Property Get Comments() As String
    Comments = mstrComments
End Property
Public Property Get MaxLength() As Long
    MaxLength = mlngMaxLength
End Property
Public Property Let MaxLength(ByVal lngValue As Long)
    mlngMaxLength = lngValue
End Property
'*******************************************************************
'Function: Get_User_Input
'Author: KDunne
'Date: 12/6/2001
'Purpose: Prompts the user for input
'*******************************************************************
Public Function Get_User_Input(Optional EnableTextBox As Boolean = True) As Boolean
    Dim blnReturn As Boolean
    
    frmComments.Caption = mstrCaption
    frmComments.txtComments.Text = mstrComments
    frmComments.Validate_Text = mblnValidate
    
    'SCR 33604
    frmComments.txtComments.Enabled = EnableTextBox
    frmComments.cmdOk.Enabled = EnableTextBox
    
    Set frmComments.Parent = Me
    
    If mlngMaxLength > 0 Then
        frmComments.SetMaxChars mlngMaxLength
    Else
        'Default to 2000.
        frmComments.SetMaxChars 2000
    End If
    
    mblnCancelled = False
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    frmComments.Show vbModal
    
    'KDunne 12/2/2002
    'SCR 29726 - If we're currently showing a modal form, don't send the Done message.
    If mblnShowingModalForm = False Then
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
    End If
    
    If mblnCancelled Then
        blnReturn = False
    Else
        blnReturn = True
    End If
    
    Get_User_Input = blnReturn
End Function

Private Sub Class_Initialize()
    mlngMaxLength = -1
    mblnShowingModalForm = False
End Sub

Private Sub IComments_SetUserComments(ByVal strComments As String)
    mstrComments = strComments
End Sub

Private Sub IComments_UserCancelled()
    mblnCancelled = True
End Sub
