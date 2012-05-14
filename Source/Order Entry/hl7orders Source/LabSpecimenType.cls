VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabSpecimenType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'For sending with an order (NOT preverify ... see SummaryType)

Private mvarTopography As String 'local copy
Private mvarName As String 'local copy
Private mvarIen As String 'local copy
Private mvarAnswerCode As String 'local copy
Private mvarAnswerText As String 'local copy
Private mvarCapColorOrContainer As String 'local copy
Private mvarTopographyIen As String 'local copy
'#If coiRequiredComments = True Then
    Private mvarCollRequiredComments As RequiredCommentsCollection

Public Property Set colRequiredComments(ByVal cCollection As RequiredCommentsCollection)
    Set mvarCollRequiredComments = cCollection
End Property

Public Property Get colRequiredComments() As RequiredCommentsCollection
    Set colRequiredComments = mvarCollRequiredComments
End Property
'#End If
Public Property Let TopographyIen(ByVal vData As String)
    mvarTopographyIen = vData
End Property

Public Property Get TopographyIen() As String
    TopographyIen = mvarTopographyIen
End Property

Public Property Let CapColorOrContainer(ByVal vData As String)
    mvarCapColorOrContainer = vData
End Property

Public Property Get CapColorOrContainer() As String
    CapColorOrContainer = mvarCapColorOrContainer
End Property

Public Property Let AnswerText(ByVal vData As String)
    mvarAnswerText = vData
End Property

Public Property Get AnswerText() As String
    AnswerText = mvarAnswerText
End Property

Public Property Let AnswerCode(ByVal vData As String)
    mvarAnswerCode = vData
End Property

Public Property Get AnswerCode() As String
    AnswerCode = mvarAnswerCode
End Property

Public Property Let IEN(ByVal vData As String)
    mvarIen = vData
End Property

Public Property Get IEN() As String
    IEN = mvarIen
End Property

Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property

Public Property Let Topography(ByVal vData As String)
    mvarTopography = vData
End Property

Public Property Get Topography() As String
    Topography = mvarTopography
End Property



Private Sub Class_Initialize()
    #If coiRequiredComments = True Then
        Set mvarCollRequiredComments = New RequiredCommentsCollection
    #End If
End Sub

Private Sub Class_Terminate()
    #If coiRequiredComments = True Then
        Set mvarCollRequiredComments = Nothing
    #End If

End Sub
