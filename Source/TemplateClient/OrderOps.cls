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
Attribute VB_Name = "OrderOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type RxItem
    szIEN As String
    szName As String
    szDefUnit As String
    szSynonym As String
    ip_op_indicator As Integer
    drug_check_disabled As Integer
    non_formulary As Integer
    Message As String
    szDisplayName As String
    Stocked As Integer
    formulary_group_ien As Integer
End Type

Public Type LabRadItem
    szDisplayName As String
    szIEN As String
    szName As String
    szSynonym As String
    szPanelFlag As String
    szSubscript As String
End Type

Public Type SPECIMEN_INFO_STRUCT
    iZoc_Count As Integer
    bSpecimenParmsReturnedFlag As Boolean        'replaces giResults
    Zoc1_SpecimenName As String
    Zoc1_szSpecimenIen As String
    Zoc2_TopographyName As String
    Zoc2_szTopoIen As String
    Zoc3_Comments As String
    Zoc4_TitleForQuestions As String
    Zoc5_iAnswersCount As Integer
    Zoc6_CapOrContainer As String
    SpecimenAnswerText As String
    SpecimenAnswerCode As String
    OrderName As String
    iZoc_Index As Integer
End Type

