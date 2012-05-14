VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSVGSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This stores the svg string for each surface on each tooth, including the svg for the event handlers

Private mvarToothID As Integer 'local copy
Private mvarSurfaceID As Integer 'local copy
Private mvarSurface As String 'local copy
Private mvarASurface As String 'local copy

Public Property Let SurfaceID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SurfaceID = 5
    mvarSurfaceID = vData
End Property


Public Property Get SurfaceID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SurfaceID
    SurfaceID = mvarSurfaceID
End Property



Public Property Let ASurface(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ASurface = 5
    mvarASurface = vData
End Property


Public Property Get ASurface() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ASurface
    ASurface = mvarASurface
End Property



Public Property Let Surface(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Surface = 5
    mvarSurface = vData
End Property


Public Property Get Surface() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Surface
    Surface = mvarSurface
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




Public Function getSurfaceDefHeader() As String
    getSurfaceDefHeader = "" 'sGrpHeader
End Function
Public Function getSurfaceDefBody() As String
    getSurfaceDefBody = mvarSurface
End Function
Public Function getASurfaceDefBody() As String
    getASurfaceDefBody = mvarASurface
End Function
Public Function getSurfaceDefFooter() As String
    getSurfaceDefFooter = "" 'sGrpFooter
End Function



