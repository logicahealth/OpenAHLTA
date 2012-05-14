VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrQuery As String
Private mstrTitle As String
Private mstrCancelCaption As String
Private mstrEnterCaption As String

Public Property Get Query() As String
   Query = mstrQuery
End Property
Public Property Let Query(strQuery As String)
   mstrQuery = strQuery
End Property

Public Function GetInput(frm As Variant) As Boolean
    
    Dim blnRet As Boolean

    Load frmQuery
    With frmQuery
      .cmdCancel.Caption = mstrCancelCaption
      .cmdOK.Caption = mstrEnterCaption
      cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_MEDCIN_APPLICATION
      .Show vbModal, frm
      cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_MEDCIN_APPLICATION
    
      If Len(.UserQuery) > 0 Then
         mstrQuery = .UserQuery
         blnRet = True
      Else
         mstrQuery = ""
         blnRet = False
      End If
    End With
   
    Unload frmQuery
    GetInput = blnRet

End Function

Public Property Get EnterCaption() As String
   EnterCaption = mstrEnterCaption
End Property
Public Property Let EnterCaption(strCaption As String)
   mstrEnterCaption = strCaption
End Property
Public Property Get CancelCaption() As String
   CancelCaption = mstrCancelCaption
End Property
Public Property Let CancelCaption(strCaption As String)
   mstrCancelCaption = strCaption
End Property

Private Sub Class_Initialize()
   mstrCancelCaption = "Cancel"
   mstrEnterCaption = "OK"
End Sub

