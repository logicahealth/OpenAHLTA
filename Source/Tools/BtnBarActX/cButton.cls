VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public key As String
Public Tag As String
Public caption As String
Public Index As Integer
Public IsDrpArrw As Boolean
Public MenuItemColorStart As Long
Public MenuItemColorEnd As Long

Private mvarMenuItemKey As String
Private Property Let MenuItemKey(vData As String)
    
    mvarMenuItemKey = vData
    
End Property

Private Property Get MenuItemKey() As String

    MenuItemKey = mvarMenuItemKey
    
End Property


