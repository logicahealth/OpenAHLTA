VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrTitle As String
Private mstrComment As String

Public Property Get Title() As String
    Title = mstrTitle
End Property
Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
End Property
Public Property Get Comment() As String
    Comment = mstrComment
End Property
Public Property Let Comment(strComment As String)
    mstrComment = strComment
End Property

Public Function GetInput() As Boolean
    
    Dim frmComment As notefrm

    On Error GoTo ErrHandler
    Set frmComment = New notefrm
    
    With frmComment
        .setText mstrComment
        .setTitle mstrTitle
        If (mstrTitle = "Finding note") Then
          frmMain2.setPadLocation frmComment
        Else
          frmMain2.setFormLocation frmComment
        End If
        'Move SHOWING MODAL FORM to cmdFreeText.Click
        'cwobjCWComm.message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_MEDCIN_APPLICATION
        .Show vbModal, frmMain2
        If .AcceptChanges Then
            'gTextNote = .getText
            mstrComment = .getText
            GetInput = True
        Else
            GetInput = False
        End If
        'cwobjCWComm.message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_MEDCIN_APPLICATION
    End With
    
    Screen.MousePointer = vbHourglass
    Unload frmComment
    GoTo CleanUpAndExit
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.description, "mMedcin.findingNote", "Template Mgmt", vbCritical)

CleanUpAndExit:
    DoEvents
    Pause 1
    Screen.MousePointer = vbDefault
End Function
