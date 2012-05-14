VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGCTx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This stores Graphical Charting info about a Tx

Private mvarDrawThisInDxViewTx As Boolean
Private mvarDrawThisInTxViewTx As Boolean
Private mvarDrawThisInDx603aViewTx As Boolean
Private mvarToothID As Integer
Private mvarSurface As String
Private mvarToothLocation As eLocation
Private mvarMenuID As Integer
Private mvarRootMenuID As Integer
Private mvarModifiedDateTime As Date
Private mvarSVGTx As String
Private mvarTxStatus As CHCSII_Dental_DXTX.TxStatusEnum
Private mvarTxStatusInBothViews As Boolean

Public Property Let TxStatus(ByVal vData As CHCSII_Dental_DXTX.TxStatusEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TxStatus = 5
    mvarTxStatus = vData
    
    If mvarTxStatusInBothViews = False Then
        If mvarTxStatus = txCompleted Or mvarTxStatus = txPreExisting Then
                mvarDrawThisInDxViewTx = True
                mvarDrawThisInTxViewTx = False
        ElseIf mvarTxStatus = txPlanned Or mvarTxStatus = txInProgress Then
                mvarDrawThisInTxViewTx = True
                mvarDrawThisInDxViewTx = False
        End If
    End If
    
    'Only draw txs in Dx603a view if the status is PreExisting or Completed
    If mvarTxStatus = txPreExisting Or mvarTxStatus = txCompleted Then
        mvarDrawThisInDx603aViewTx = True
    Else
        mvarDrawThisInDx603aViewTx = False
    End If
    
End Property


Public Property Let TxStatusInBothViews(vData As Boolean)
    
    
    'Show this Tx in both views,
    'can only show in both views if the status is completed or preexisting
    If vData = True And mvarTxStatus = txCompleted Or mvarTxStatus = txPreExisting Then
        mvarTxStatusInBothViews = True
        mvarDrawThisInDxViewTx = True
        mvarDrawThisInTxViewTx = True
    Else
        mvarTxStatusInBothViews = False
    End If
    
End Property

Public Property Get TxStatus() As CHCSII_Dental_DXTX.TxStatusEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TxStatus
    TxStatus = mvarTxStatus
End Property



Public Property Let SVGTx(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SVGTx = 5
    mvarSVGTx = vData
End Property


Public Property Get SVGTx() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SVGTx
    SVGTx = mvarSVGTx
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



Public Property Let DrawThisInTxViewTx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawThisInTxViewTx = 5
    mvarDrawThisInTxViewTx = vData
    
End Property


Public Property Get DrawThisInTxViewTx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawThisInTxViewTx
    DrawThisInTxViewTx = mvarDrawThisInTxViewTx
End Property



Public Property Let DrawThisInDxViewTx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawThisInDxViewTx = 5
    mvarDrawThisInDxViewTx = vData
End Property


Public Property Get DrawThisInDxViewTx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawThisInDxViewTx
    DrawThisInDxViewTx = mvarDrawThisInDxViewTx
End Property

Public Property Let DrawThisInDx603aViewTx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawThisInDxViewTx = 5
    mvarDrawThisInDx603aViewTx = vData
End Property


Public Property Get DrawThisInDx603aViewTx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawThisInDxViewTx
    DrawThisInDx603aViewTx = mvarDrawThisInDx603aViewTx
End Property


