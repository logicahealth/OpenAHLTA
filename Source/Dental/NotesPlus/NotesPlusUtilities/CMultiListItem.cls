VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultiListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrKey As String
Private mstrDesc As String
Private mblnChecked As Boolean
Private mblnPrimary As Boolean
Private mblnSecondary As Boolean

Public Property Get Key() As String
    Key = mstrKey
End Property

Public Property Get Description() As String
    Description = mstrDesc
End Property

Public Property Get Checked() As Boolean
    Checked = mblnChecked
End Property

Friend Property Let Key(ByVal strKey As String)
    mstrKey = strKey
End Property
Friend Property Let Description(ByVal strDesc As String)
    mstrDesc = strDesc
End Property
Friend Property Let Checked(ByVal blnChecked As Boolean)
    mblnChecked = blnChecked
End Property
Friend Property Let Primary(ByVal blnChecked As Boolean)
   mblnPrimary = blnChecked
End Property
Public Property Get Primary() As Boolean
    Primary = mblnPrimary
End Property
Friend Property Let Secondary(ByVal blnChecked As Boolean)
    mblnSecondary = blnChecked
End Property
Public Property Get Secondary() As Boolean
    Secondary = mblnSecondary
End Property
