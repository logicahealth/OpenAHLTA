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
Attribute VB_Name = "thr_Facility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moSites     As GEMS_Sites
Private moCurrent   As GEMS_Site
Private mvaridx     As Long
Private mvarBOF     As Boolean
Private mvarEOF     As Boolean

Implements IFacility

Friend Property Set GemsFacility(GemsFacilities As GEMS_Sites)

    Set moSites = GemsFacilities
    mvarBOF = True
    mvaridx = 0
    
    If moSites.Count = 0 Then
        mvarEOF = True
    Else
        mvaridx = 1
        FindCurrentIndex
        mvarEOF = False
    End If

End Property
Private Sub FindCurrentIndex()
    Dim lpos    As Long
    Dim v       As GEMS_Site

    lpos = 1
    For Each v In moSites

        lpos = lpos + 1
        If lpos > mvaridx Then
            Set moCurrent = v
            Exit For
        End If
    Next v

End Sub
Private Property Get IFacility_BOF() As Boolean

    IFacility_BOF = mvarBOF

End Property


Private Property Get IFacility_EOF() As Boolean

    IFacility_EOF = mvarEOF

End Property

Private Sub IFacility_GotoFirst()

    mvaridx = 1
    If moSites.Count = 0 Then
      mvarBOF = True
      mvarEOF = True
    ElseIf mvaridx <= moSites.Count Then
      mvarBOF = False
      mvarEOF = False
      FindCurrentIndex
    Else
      mvarEOF = True
      mvaridx = 0
    End If
    
End Sub

Private Sub IFacility_GotoLast()
    Dim v As Variant

    mvaridx = 0
    For Each v In moSites
        mvaridx = mvaridx + 1
    Next v
    FindCurrentIndex

End Sub

Private Sub IFacility_GotoNext()

   mvaridx = mvaridx + 1
    If mvaridx > moSites.Count Then
        mvarEOF = True
        Debug.Assert mvaridx > moSites.Count
    End If
    FindCurrentIndex
    
End Sub

Private Sub IFacility_GotoPrevious()

    If mvaridx = 1 Then
        mvarBOF = True
    ElseIf mvaridx <= 0 Then
        Debug.Assert mvaridx >= 0
        Exit Sub
    End If
    
    mvaridx = mvaridx - 1
    
    Debug.Assert mvaridx > 0
    FindCurrentIndex

End Sub

Private Property Let IFacility_MedRecNum(ByVal RHS As String)

    moCurrent.MedRecNum = RHS

End Property

Private Property Get IFacility_MedRecNum() As String

    IFacility_MedRecNum = moCurrent.MedRecNum

End Property

Private Property Let IFacility_Name(ByVal RHS As String)

    moCurrent.FacilityNCID = RHS

End Property

Private Property Get IFacility_Name() As String

    IFacility_Name = moCurrent.FacilityNCID

End Property

Private Property Let IFacility_RadiologyNum(ByVal RHS As String)

    moCurrent.RadiologyNum = RHS

End Property

Private Property Get IFacility_RadiologyNum() As String

    IFacility_RadiologyNum = moCurrent.RadiologyNum

End Property

Private Property Get IFacility_StoreTime() As Date

End Property

Private Property Let IFacility_StoreTime(ByVal RHS As Date)

End Property

Private Property Let IFacility_Type(ByVal RHS As String)

    If mvaridx = 0 Then
        IFacility_GotoLast
    End If

    If mvaridx > moSites.Count Or moSites.Count = 0 Then
        Dim gemsSite As GEMS_Site
        Set gemsSite = New GEMS_Site
    
        gemsSite.UpdateFlag = "N"
        moSites.AddSite gemsSite
        mvaridx = moSites.Count
        FindCurrentIndex
    End If

    moCurrent.FacilityNCID = RHS
    'Type 0 means delete, 3M documentation states
    If Val(RHS) = 0 Then
        moCurrent.UpdateFlag = "D"
    End If
    
End Property

Private Property Get IFacility_Type() As String

    IFacility_Type = moCurrent.FacilityNCID
    
End Property
