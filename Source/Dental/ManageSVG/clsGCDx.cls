VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGCDx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This stores Graphical Charting info about a Dx

Private mvarDrawThisInDxViewDx As Boolean
Private mvarDrawThisInTxViewDx As Boolean
Private mvarDrawThisInDx603aViewDx As Boolean
Private mvarToothID As Integer
Private mvarSurface As String
Private mvarToothLocation As eLocation
Private mvarMenuID As Integer
Private mvarRootMenuID As Integer
Private mvarModifiedDateTime As Date
Private mvarSVGDx As String
Private mvarDRC As Integer
Private mvarDxStatus As CHCSII_Dental_DXTX.DxStatusEnum
Public Property Let DxStatus(ByVal vData As CHCSII_Dental_DXTX.DxStatusEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DxStatus = 5
    mvarDxStatus = vData

    If (mvarDxStatus = dxActiveNew) Or (mvarDxStatus = dxActiveOngoing) Or (mvarDxStatus = dxActiveProvisional) Then
        mvarDrawThisInDxViewDx = True
        mvarDrawThisInTxViewDx = True
    End If

End Property


Public Property Get DxStatus() As CHCSII_Dental_DXTX.DxStatusEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DxStatus
    DxStatus = mvarDxStatus
End Property



Public Property Let DRC(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DRC = 5
    mvarDRC = vData
End Property


Public Property Get DRC() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DRC
    DRC = mvarDRC
End Property



Public Property Let SVGDx(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SVGDx = 5
    mvarSVGDx = vData
End Property


Public Property Get SVGDx() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SVGDx
    SVGDx = mvarSVGDx
End Property



Public Property Let ModifiedDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ModifiedDateTime = 5
    mvarModifiedDateTime = vData
End Property


Public Property Get ModifiedDateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ModifiedDateTime
    ModifiedDateTime = mvarModifiedDateTime
End Property



Public Property Let RootMenuID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootMenuID = 5
    mvarRootMenuID = vData
End Property


Public Property Get RootMenuID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootMenuID
    RootMenuID = mvarRootMenuID
End Property



Public Property Let MenuID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MenuID = 5
    mvarMenuID = vData
End Property


Public Property Get MenuID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuID
    MenuID = mvarMenuID
End Property



Public Property Let ToothLocation(ByVal vData As eLocation)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ToothLocation = Form1
    mvarToothLocation = vData
End Property


Public Property Get ToothLocation() As eLocation
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothLocation
    ToothLocation = mvarToothLocation
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



Public Property Let DrawThisInTxViewDx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawThisInTxViewDx = 5
    mvarDrawThisInTxViewDx = vData
End Property


Public Property Get DrawThisInTxViewDx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawThisInTxViewDx
    DrawThisInTxViewDx = mvarDrawThisInTxViewDx
End Property



Public Property Let DrawThisInDxViewDx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawThisInDxViewDx = 5
    mvarDrawThisInDxViewDx = vData
End Property


Public Property Get DrawThisInDxViewDx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawThisInDxViewDx
    DrawThisInDxViewDx = mvarDrawThisInDxViewDx
End Property

Public Property Let DrawThisInDx603aViewDx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawThisInDxViewDx = 5
    mvarDrawThisInDx603aViewDx = vData
End Property


Public Property Get DrawThisInDx603aViewDx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawThisInDxViewDx

    'Never draw dxs in Dx603a View
    DrawThisInDx603aViewDx = False

    'DrawThisInDx603aViewDx = mvarDrawThisInDx603aViewDx
    
End Property


