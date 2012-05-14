'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarPosition As Integer
Private mvarToothNum As Long
Private mvarSeen As Boolean
Private mvarStatus As ePositionStatus
Private mvarSelected As Boolean
Private mvarToothErupted As Boolean

Private mbPropertyChanged As Boolean

Public Enum ePositionStatus
    'eNone = 0
    ePrimary = 1
    eSecondary = 2
End Enum

Friend Property Get PropertyChanged() As Boolean
    
    PropertyChanged = mbPropertyChanged
    
End Property



Friend Property Let PropertyChanged(vData As Boolean)
    
    mbPropertyChanged = vData
    
End Property

Public Property Let Selected(vData As Boolean)

    mvarSelected = vData
    
End Property

Public Property Get Selected() As Boolean

    Selected = mvarSelected
    
End Property


Public Property Let Status(ByVal vData As ePositionStatus)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Status = Form1
    
    If vData = 0 Then
        'This is for backwards compatibility for database entries.
        'If the Status is 0 then it is not erupted
        Me.ToothErupted = False
        Exit Property
    End If
    
    'If I change the status, then I change the tooth number because the position is constant
    mvarStatus = vData
    
    If mvarStatus = ePositionStatus.eSecondary Then
        mvarToothNum = mvarPosition
    ElseIf mvarStatus = ePositionStatus.ePrimary Then
        'Convert the position to it's primary tooth number
        mvarToothNum = ConvPosToPrimaryNum(mvarPosition)
    End If

    mbPropertyChanged = True

End Property


Public Property Get Status() As ePositionStatus
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Friend Property Let Position(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Position = 5
    mvarPosition = vData
    
    If mvarPosition = 0 Then
        mvarToothNum = 0
    ElseIf mvarStatus = ePositionStatus.eSecondary Then
        mvarToothNum = mvarPosition
    ElseIf mvarStatus = ePositionStatus.ePrimary Then
        'Convert the position to it's primary tooth number
        mvarToothNum = ConvPosToPrimaryNum(mvarPosition)
    End If
    
    mbPropertyChanged = True
    
End Property


Public Property Get Position() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Position
    Position = mvarPosition
End Property



Public Property Let Seen(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Seen = 5
    
    If mvarSeen = vData Then
        Exit Property
    End If
    
    mvarSeen = vData
    
    mbPropertyChanged = True
    
End Property


Public Property Get Seen() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Seen
    Seen = mvarSeen
End Property



Public Property Let ToothErupted(vData As Boolean)

    mvarToothErupted = vData
    
    'If a tooth has not erupted in this position, then show the position as seen
    If mvarToothErupted = False Then
        mvarSeen = True
    End If
    
    mbPropertyChanged = True
    
End Property

Public Property Get ToothErupted() As Boolean

    ToothErupted = mvarToothErupted
    
End Property

Private Property Let ToothNum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothNum = 5

    'Don't let the user change the tooth number, this is done automatically by changing the status
    'and is based on the position.
    mvarToothNum = vData
    
    mbPropertyChanged = True

End Property


Public Property Get ToothNum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothNum
    ToothNum = mvarToothNum
End Property





