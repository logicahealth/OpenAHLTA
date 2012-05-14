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
Attribute VB_Name = "CytologyGyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarBirthControlPills As hlCHCS_Boolean
Private mvarIUD As hlCHCS_Boolean 'local copy
Private mvarPostMenopausal As hlCHCS_Boolean 'local copy
Private mvarHysterectomy As hlCHCS_Boolean 'local copy
Private mvarHormoneTherapy As hlCHCS_Boolean 'local copy
Private mvarLMPDate As String 'local copy
Private mvarWeeksPregnant As String 'local copy
Private mvarWeeksPostPartum As String 'local copy
Private mvarRadiationTherapyDate As String 'local copy
Private mvarCytotoxicTherapyDate As String 'local copy
Private mvarPreviousCytologyDiagnosis As String 'local copy
Private mvarSpecimenCollectiondate As String
Private mvarTopoIEN As String
Private mvarTopoName As String
Private mvarComments As String
Public Property Let TopoName(ByVal vData As String)
    mvarTopoName = vData
End Property
Public Property Get TopoName() As String
    TopoName = mvarTopoName
End Property
Public Property Let Comments(ByVal vData As String)
    mvarComments = vData
End Property
Public Property Get Comments() As String
    Comments = mvarComments
End Property


Public Property Let TopoIEN(ByVal vData As String)
    mvarTopoIEN = vData
End Property


Public Property Get TopoIEN() As String
    TopoIEN = mvarTopoIEN
End Property
Public Property Let SpecimenCollectionDate(ByVal vData As String)
    mvarSpecimenCollectiondate = vData
End Property


Public Property Get SpecimenCollectionDate() As String
    SpecimenCollectionDate = mvarSpecimenCollectiondate
End Property
Public Property Let PreviousCytologyDiagnosis(ByVal vData As String)
    mvarPreviousCytologyDiagnosis = vData
End Property


Public Property Get PreviousCytologyDiagnosis() As String
    PreviousCytologyDiagnosis = mvarPreviousCytologyDiagnosis
End Property
Public Property Let CytotoxicTherapyDate(ByVal vData As String)
    mvarCytotoxicTherapyDate = vData
End Property


Public Property Get CytotoxicTherapyDate() As String
    CytotoxicTherapyDate = mvarCytotoxicTherapyDate
End Property
Public Property Let RadiationTherapyDate(ByVal vData As String)
    mvarRadiationTherapyDate = vData
End Property
Public Property Get RadiationTherapyDate() As String
    RadiationTherapyDate = mvarRadiationTherapyDate
End Property

Public Property Let WeeksPostPartum(ByVal vData As String)
    mvarWeeksPostPartum = vData
End Property


Public Property Get WeeksPostPartum() As String
    WeeksPostPartum = mvarWeeksPostPartum
End Property
Public Property Let WeeksPregnant(ByVal vData As String)
    mvarWeeksPregnant = vData
End Property

Public Property Get WeeksPregnant() As String
    WeeksPregnant = mvarWeeksPregnant
End Property
Public Property Let LMPDate(ByVal vData As String)
    mvarLMPDate = vData
End Property


Public Property Get LMPDate() As String
    LMPDate = mvarLMPDate
End Property



Public Property Let HormoneTherapy(ByVal vData As hlCHCS_Boolean)
    mvarHormoneTherapy = vData
End Property


Public Property Get HormoneTherapy() As hlCHCS_Boolean
   HormoneTherapy = mvarHormoneTherapy
End Property



Public Property Let Hysterectomy(ByVal vData As hlCHCS_Boolean)
    mvarHysterectomy = vData
End Property


Public Property Get Hysterectomy() As hlCHCS_Boolean
   Hysterectomy = mvarHysterectomy
End Property



Public Property Let PostMenopausal(ByVal vData As hlCHCS_Boolean)
    mvarPostMenopausal = vData
End Property


Public Property Get PostMenopausal() As hlCHCS_Boolean
   PostMenopausal = mvarPostMenopausal
End Property



Public Property Let IUD(ByVal vData As hlCHCS_Boolean)
    mvarIUD = vData
End Property


Public Property Get IUD() As hlCHCS_Boolean
    IUD = mvarIUD
End Property



Public Property Let BirthControlPills(ByVal vData As hlCHCS_Boolean)
    mvarBirthControlPills = vData
End Property


Public Property Get BirthControlPills() As hlCHCS_Boolean
    BirthControlPills = mvarBirthControlPills
End Property



Private Sub Class_Initialize()
    mvarBirthControlPills = hlNull
    mvarHormoneTherapy = hlNull
    mvarIUD = hlNull
    mvarPostMenopausal = hlNull
    mvarHysterectomy = hlNull
    
    
End Sub

