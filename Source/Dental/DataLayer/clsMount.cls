VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRads As colReadOnlyRads
Private mvarMountID As Long
Private mvarPatienUnitID As Long
Private mvarDirtyFlag As DirtyFlagEnum
Friend Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DirtyFlag = Form1
    mvarDirtyFlag = vData
End Property


Friend Property Get DirtyFlag() As DirtyFlagEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirtyFlag
    DirtyFlag = mvarDirtyFlag
End Property



Friend Property Let PatienUnitID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatienUnitID = 5
    mvarPatienUnitID = vData
End Property


Friend Property Get PatienUnitID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatienUnitID
    PatienUnitID = mvarPatienUnitID
End Property



Public Property Let MountID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MountID = 5
    mvarMountID = vData
End Property


Public Property Get MountID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MountID
    MountID = mvarMountID
End Property



Public Property Set Rads(ByVal vData As colReadOnlyRads)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Rads = Form1
    Set mvarRads = vData
End Property


Public Property Get Rads() As colReadOnlyRads
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rads

    If mvarRads Is Nothing Then
        Set mvarRads = New colReadOnlyRads
    End If
    
    Set Rads = mvarRads
    
End Property



Private Sub Class_Terminate()

    Set mvarRads = Nothing
    
End Sub


