VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mintPage As Integer
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long
Private mlngSnoID As String
Private mlngFlags As Long
Private mstrPrefix As String
Private mstrItemData As String
Private mstrDescription As String

Public Property Get Page() As Integer
   Page = mintPage
End Property
Public Property Let Page(intPage As Integer)
   mintPage = intPage
End Property
Public Property Get Left() As Long
   Left = mlngLeft
End Property
Public Property Let Left(lngLeft As Long)
   mlngLeft = lngLeft
End Property
Public Property Get Top() As Long
   Top = mlngTop
End Property
Public Property Let Top(lngTop As Long)
   mlngTop = lngTop
End Property
Public Property Get Right() As Long
   Right = mlngRight
End Property
Public Property Let Right(lngRight As Long)
   mlngRight = lngRight
End Property
Public Property Get Bottom() As Long
   Bottom = mlngBottom
End Property
Public Property Let Bottom(lngBottom As Long)
   mlngBottom = lngBottom
End Property
Public Property Get SnoID() As Long
   SnoID = mlngSnoID
End Property
Public Property Let SnoID(lngId As Long)
   mlngSnoID = lngId
End Property
Public Property Get Flags() As Long
   Flags = mlngFlags
End Property
Public Property Let Flags(lngFlags As Long)
   mlngFlags = lngFlags
End Property
Public Property Get Prefix() As String
   Prefix = mstrPrefix
End Property
Public Property Let Prefix(strPrefix As String)
   mstrPrefix = strPrefix
End Property
Public Property Get ItemData() As String
   ItemData = mstrItemData
End Property
Public Property Let ItemData(strItemData As String)
   mstrItemData = strItemData
End Property
Public Property Get Description() As String
   Description = mstrDescription
End Property
Public Property Let Description(strDesc As String)
   mstrDescription = strDesc
End Property
