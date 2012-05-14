VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)

Private mvarShowNewOrders As Boolean 'local copy
Private mvarShowSystemOrders As Boolean 'local copy
Private mvarOrderTypeCode As String 'local copy
Private mvarSortBy As String 'local copy
Private mvarSortOrder As String 'local copy
Private mvarOrderCount As String 'local copy

Public Property Let OrderCount(ByVal vData As String)
    mvarOrderCount = vData
End Property

Public Property Get OrderCount() As String
    OrderCount = mvarOrderCount
End Property

Public Property Let SortOrder(ByVal vData As String)
    mvarSortOrder = vData
End Property

Public Property Get SortOrder() As String
    SortOrder = mvarSortOrder
End Property

Public Property Let SortBy(ByVal vData As String)
    mvarSortBy = vData
End Property

Public Property Get SortBy() As String
    SortBy = mvarSortBy
End Property

Public Property Let OrderTypeCode(ByVal vData As String)
    mvarOrderTypeCode = vData
End Property

Public Property Get OrderTypeCode() As String
    OrderTypeCode = mvarOrderTypeCode
End Property

Public Property Let ShowSystemOrders(ByVal vData As Boolean)
    mvarShowSystemOrders = vData
End Property

Public Property Get ShowSystemOrders() As Boolean
    ShowSystemOrders = mvarShowSystemOrders
End Property

Public Property Let ShowPreactiveOrders(ByVal vData As Boolean)
    mvarShowNewOrders = vData
End Property

Public Property Get ShowPreactiveOrders() As Boolean
    ShowPreactiveOrders = mvarShowNewOrders
End Property
