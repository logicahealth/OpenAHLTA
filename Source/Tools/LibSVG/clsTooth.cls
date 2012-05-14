VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSVGToothDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This stores the svg for each tooth

Private mvarToothID As Integer
Private mvarCrown As String
Private mvarCrownMD As String
Private mvarFooter As String
Private mvarHeader As String
Private mvarRect As String
Private mvarRoots As String
Private mvarToothName As String
Private mvarWholeTooth As String
Public Property Let WholeTooth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WholeTooth = 5
    mvarWholeTooth = vData
End Property


Public Property Get WholeTooth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WholeTooth
    WholeTooth = mvarWholeTooth
End Property



Public Property Let ToothName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothName = 5
    mvarToothName = vData
End Property


Public Property Get ToothName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothName
    ToothName = mvarToothName
End Property



Public Property Let Roots(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Roots = 5
    mvarRoots = vData
End Property


Public Property Get Roots() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Roots
    Roots = mvarRoots
End Property



Public Property Let Rect(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Rect = 5
    mvarRect = vData
End Property


Public Property Get Rect() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Rect
    Rect = mvarRect
End Property



Public Property Let Header(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Header = 5
    mvarHeader = vData
End Property


Public Property Get Header() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Header
    Header = mvarHeader
End Property



Public Property Let Footer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Footer = 5
    mvarFooter = vData
End Property


Public Property Get Footer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Footer
    Footer = mvarFooter
End Property



Public Property Let CrownMD(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CrownMD = 5
    mvarCrownMD = vData
End Property


Public Property Get CrownMD() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CrownMD
    CrownMD = mvarCrownMD
End Property



Public Property Let Crown(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Crown = 5
    mvarCrown = vData
End Property


Public Property Get Crown() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Crown
    Crown = mvarCrown
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



Public Function getToothDefHeader(Optional ToothMode As eToothMode = eBigTooth) As String

    If ToothMode = eBigTooth Then
        getToothDefHeader = mvarHeader & " " & mvarRect
    ElseIf ToothMode = eFullMouth Then
        getToothDefHeader = mvarHeader & " " & mvarRect & " " & mvarToothName
    End If
    
End Function
Public Function getToothDefBody() As String
    getToothDefBody = mvarCrownMD & " " & mvarRoots
End Function
Public Function getToothCrownBody() As String
    getToothCrownBody = mvarCrownMD
End Function
Public Function getToothRootBody() As String
    getToothRootBody = mvarRoots
End Function
Public Function getToothDefFooter() As String
    getToothDefFooter = mvarFooter
End Function

