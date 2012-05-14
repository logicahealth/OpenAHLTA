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
Attribute VB_Name = "EncEandMCoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPatientUnitNumber       As String
Private mvarEncounterNumber         As String
Private mvarEncounterFacilityNCID   As String
Private mvarEandMNCID               As String
Private mvarEandMReviewed           As String
Private mvarEandMReviewedBy         As String
Private mvarEandMReviewedDate       As String
Private mvarCurrentRTF              As String
Private mvarCodeReview              As String


Public Property Get CriteriaPatientUnitNumber() As String:       CriteriaPatientUnitNumber = mvarPatientUnitNumber: End Property
Public Property Let CriteriaPatientUnitNumber(RHV As String):    mvarPatientUnitNumber = RHV:                       End Property

Public Property Get CriteriaEncounterNumber() As String:         CriteriaEncounterNumber = mvarEncounterNumber:     End Property
Public Property Let CriteriaEncounterNumber(RHV As String):      mvarEncounterNumber = RHV:                         End Property

Public Property Get CriteriaEncounterFacilityNCID() As String:   CriteriaEncounterFacilityNCID = mvarEncounterFacilityNCID:  End Property
Public Property Let CriteriaEncounterFacilityNCID(RHV As String): mvarEncounterFacilityNCID = RHV:                  End Property

Public Property Get EandMNCID() As String:               EandMNCID = mvarEandMNCID:                  End Property
Public Property Let EandMNCID(RHV As String):            mvarEandMNCID = RHV:                        End Property

Public Property Get EandMReviewed() As String:           EandMReviewed = mvarEandMReviewed:          End Property
Public Property Let EandMReviewed(RHV As String):        mvarEandMReviewed = RHV:                    End Property

Public Property Get EandMReviewedDate() As String:       EandMReviewedDate = mvarEandMReviewedDate:  End Property
Public Property Let EandMReviewedDate(RHV As String):    mvarEandMReviewedDate = RHV:                End Property

Public Property Get EandMReviewedBy() As String:         EandMReviewedBy = mvarEandMReviewedBy:      End Property
Public Property Let EandMReviewedBy(RHV As String):      mvarEandMReviewedBy = RHV:                  End Property

Public Property Get CurrentRTF() As String:              CurrentRTF = mvarCurrentRTF:            End Property
Public Property Let CurrentRTF(RHV As String):           mvarCurrentRTF = RHV:                   End Property

Public Property Let CodeReview(ByVal vData As String): mvarCodeReview = vData:            End Property
Public Property Get CodeReview() As String:            CodeReview = mvarCodeReview:       End Property

