VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeSpecimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Built from Preverify ... for collections building ...

Private mvarSpecimenType As String 'local copy
Private mvarTopography As String 'local copy
Private mvarComments As String 'local copy
Private mvarTitle As String 'local copy
Private mvarCapColor As String 'local copy
Private mvarSpecimenTypeIen As String 'local copy
Private mvarTopographyIen As String 'local copy
      ''Private mvarAnswer As Answers 'local copy
Private mvarAnswer As Collection 'local copy
Private mvarCollRequiredComments As RequiredCommentsCollection
'Answer collection keeps the ZOC-5 list of answers, as answers objects
' that are provided in the Zp Ack.
'The question for those answers comes from ZOC-4, and is stored in the
'titleforquestions property. (HIV EIA is a good example).
Friend Property Set Answer(ByVal cCollection As Collection)
    Set mvarAnswer = cCollection
End Property

Public Property Get Answer() As Collection
    Set Answer = mvarAnswer
End Property
Public Property Set colRequiredComments(ByVal cCollection As RequiredCommentsCollection)
    Set mvarCollRequiredComments = cCollection
End Property

Public Property Get colRequiredComments() As RequiredCommentsCollection
    Set colRequiredComments = mvarCollRequiredComments
End Property
Public Property Let TopographyIen(ByVal vData As String)
    mvarTopographyIen = vData
End Property

Public Property Get TopographyIen() As String
    TopographyIen = mvarTopographyIen
End Property

Public Property Let Ien(ByVal vData As String)
    mvarSpecimenTypeIen = vData
End Property

Public Property Get Ien() As String
    Ien = mvarSpecimenTypeIen
End Property

Public Property Let CapColorOrContainer(ByVal vData As String)
    mvarCapColor = vData
End Property

Public Property Get CapColorOrContainer() As String
    CapColorOrContainer = mvarCapColor
End Property

Public Property Let TitleForQuestions(ByVal vData As String)
    mvarTitle = vData
End Property

Public Property Get TitleForQuestions() As String
    TitleForQuestions = mvarTitle
End Property

Public Property Let Comments(ByVal vData As String)
    mvarComments = vData
End Property

Public Property Get Comments() As String
    Comments = mvarComments
End Property

Public Property Let Topography(ByVal vData As String)
    mvarTopography = vData
End Property

Public Property Get Topography() As String
    Topography = mvarTopography
End Property

Public Property Let Name(ByVal vData As String)
    mvarSpecimenType = vData
End Property

Public Property Get Name() As String
    Name = mvarSpecimenType
End Property



