VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This stores a mapping of what tooth has what surfaces for crown, roots, etc.

Private mvarCrown_XY As String
Private mvarToothID As Integer
Private mvarCrownSurfaces As String
Private mvarOcclusalSurfaces As String
Private mvarRoot_XY As String
Private mvarRootSurfaces As String
Public Property Let RootSurfaces(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootSurfaces = 5
    mvarRootSurfaces = vData
End Property


Public Property Get RootSurfaces() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootSurfaces
    RootSurfaces = mvarRootSurfaces
End Property



Public Property Let Root_XY(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Root_XY = 5
    mvarRoot_XY = vData
End Property


Public Property Get Root_XY() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Root_XY
    Root_XY = mvarRoot_XY
End Property



Public Property Let OcclusalSurfaces(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OcclusalSurfaces = 5
    mvarOcclusalSurfaces = vData
End Property


Public Property Get OcclusalSurfaces() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OcclusalSurfaces
    OcclusalSurfaces = mvarOcclusalSurfaces
End Property



Public Property Let CrownSurfaces(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CrownSurfaces = 5
    mvarCrownSurfaces = vData
End Property


Public Property Get CrownSurfaces() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CrownSurfaces
    CrownSurfaces = mvarCrownSurfaces
End Property



Public Property Let ToothID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarToothID = vData
End Property


Public Property Get ToothID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    ToothID = mvarToothID
End Property



Public Property Let Crown_XY(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Crown_XY = 5
    mvarCrown_XY = vData
End Property


Public Property Get Crown_XY() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Crown_XY
    Crown_XY = mvarCrown_XY
End Property



