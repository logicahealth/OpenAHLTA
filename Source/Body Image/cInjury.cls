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
Attribute VB_Name = "cInjury"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mPart   As ePart
Private mSideLR As eSide
Private mSideFB As eSide
Private sInjury As String
Private mDesc   As String
Public FromBI   As Boolean
Public TypeID   As eType
Public DataID   As Long
Public Sno As cBISno
Public RFVcomment  As String
Public NewFollowup As String
Public NotSaved As Boolean

Public RFVDataID As String

Public lblIndex As Integer
Public lblLeft  As Long
Public lblTop   As Long

Public Enum ePart
    epUpperArm = 1
    epLowerArm = 2
    epUpperLeg = 3
    epLowerLeg = 4
    epFoot = 5
    epPelvis = 6
    epAbdomen = 7
    epChest = 8
    epSholder = 9
    epNeck = 10
    epHead = 11
    epHand = 12
'<< scr 39465 pain in back not reported properly
    epUpper = 8
    epLower = 7
End Enum

Public Enum eSide
    esLeft = 0
    esRight = 1
    esFront = 2
    esBack = 3
End Enum


Public Property Get Part() As ePart
    Part = mPart
End Property

Public Property Get SideLR() As eSide: SideLR = mSideLR: End Property
Public Property Get SideFB() As eSide: SideFB = mSideFB: End Property
Public Property Get Desc() As String: Desc = mDesc: End Property
Public Property Let Desc(RHS As String)
    mDesc = RHS
    Call ParseDesc
End Property
Public Property Get Injury() As String: Injury = sInjury: End Property

Public Property Let Injury(RHS As String)
    sInjury = RHS
    Select Case LCase$(RHS)
    Case "gunshot":    TypeID = etGunshot
    Case "fracture":   TypeID = etFracture
    Case "laceration": TypeID = etlaceration
    Case "amputation": TypeID = etAmputation
    Case "impaled":    TypeID = etImpaled
    Case "burn":       TypeID = etBurn
    Case "tourniquet":  TypeID = etTourniquet
    Case "bruise":     TypeID = etBruising
    Case "swelling":   TypeID = etSwelling
    Case "pain":       TypeID = etPain
    End Select
End Property

Public Property Let Part(RHS As ePart)
    mPart = RHS
End Property
Public Property Let SideLR(RHS As eSide)
    mSideLR = RHS
End Property
Public Property Let SideFB(RHS As eSide)
    mSideFB = RHS
End Property

Private Sub FillDesc()
    mDesc = sInjury & "- "
    mDesc = mDesc & SideEnumToText(mSideLR) & " "
    mDesc = mDesc & PartEnumToText(mPart) & ", "
    mDesc = mDesc & SideEnumToText(mSideFB)
End Sub

Private Sub ParseDesc()
    Dim sDesc As String
    Dim iPosn As Integer
    
    sDesc = Trim$(mDesc)
    iPosn = InStr(sDesc, ",")
    If iPosn > 0 Then
        If Mid$(LCase$(sDesc), iPosn + 2, 5) = "front" Then
            mSideFB = esFront
        ElseIf Mid$(LCase$(sDesc), iPosn + 2, 4) = "back" Then
            mSideFB = esBack
        End If
        sDesc = Left$(sDesc, iPosn - 1)
'<<< scr 39465 pain in back not reported properly
    ElseIf Mid$(LCase$(sDesc), 1, 5) = "lower" _
        Or Mid$(LCase$(sDesc), 1, 5) = "upper" Then
        mSideFB = esBack
    End If
    
    If Left$(LCase$(sDesc), 4) = "left" Then
        mSideLR = esLeft
        sDesc = Trim$(Mid$(sDesc, 6))  '- drop the first word
    ElseIf Left$(LCase$(sDesc), 5) = "right" Then
        mSideLR = esRight
        sDesc = Trim$(Mid$(sDesc, 7))
    End If
    mPart = PartTextToEnum(sDesc)
    Set Sno = FindSno(Me)
End Sub

Private Sub Class_Initialize()
    mSideFB = esFront
    NewFollowup = "New"
    Set Sno = New cBISno
End Sub


