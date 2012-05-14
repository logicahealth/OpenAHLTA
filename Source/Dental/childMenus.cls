VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "childMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:48:17

'-------------------------------
'This is the individual representation of a child ole module.
'It contains information needed to identify and control those childres.
'The child module itself is the mobjModule object
'The rest of this object mimics the core info array.
'-------------------------------
Private mlMenuItem  As Long


Public Property Let menuItem(RHV As Long)

   mlMenuItem = RHV

End Property

Public Property Get menuItem() As Long

   menuItem = mlMenuItem

End Property

