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
Attribute VB_Name = "thr_Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moPhones    As GEMS_Phones
Private moCurrent   As GEMS_Phone
Private mvaridx     As Long
Private mvarBOF     As Boolean
Private mvarEOF     As Boolean

Implements IPhone

Friend Property Set GemsPhone(gemsPhones As GEMS_Phones)

    Set moPhones = gemsPhones
    mvarBOF = True
    mvaridx = 0
    
    If moPhones.Count = 0 Then
        mvarEOF = True
    Else
        mvarEOF = False
    End If

End Property
Private Sub FindCurrentIndex()
    Dim lpos    As Long
    Dim v       As GEMS_Phone

    lpos = 1
    For Each v In moPhones

        lpos = lpos + 1
        If lpos > mvaridx Then
            Set moCurrent = v
            Exit For
        End If
    Next v

End Sub

Private Property Let IPhone_AreaCode(ByVal RHS As String)

    moCurrent.AreaCode = RHS

End Property

Private Property Get IPhone_AreaCode() As String

    IPhone_AreaCode = moCurrent.AreaCode

End Property


Private Property Get IPhone_BOF() As Boolean

    IPhone_BOF = mvarBOF

End Property

Private Property Let IPhone_CountryCode(ByVal RHS As String)

    moCurrent.CountryCode = RHS

End Property

Private Property Get IPhone_CountryCode() As String

    IPhone_CountryCode = moCurrent.CountryCode

End Property

Private Property Get IPhone_EOF() As Boolean

    IPhone_EOF = mvarEOF

End Property

Private Property Let IPhone_Extension(ByVal RHS As String)

    moCurrent.Extension = RHS

End Property

Private Property Get IPhone_Extension() As String

    IPhone_Extension = moCurrent.Extension

End Property

'routine copied from GEMS_Addresses
Private Sub IPhone_GotoFirst()
    
    mvaridx = 1
    
    If moPhones.Count = 0 Then
        mvarBOF = True
        mvarEOF = True
    ElseIf mvaridx <= moPhones.Count Then
      mvarBOF = False
      mvarEOF = False
      FindCurrentIndex
    Else
      mvarEOF = True
      mvaridx = 0
    End If

End Sub

Private Sub IPhone_GotoLast()
    Dim v As Variant

    mvaridx = 0
    For Each v In moPhones
        mvaridx = mvaridx + 1
    Next v

    FindCurrentIndex

End Sub

Private Sub IPhone_GotoNext()

    mvaridx = mvaridx + 1
    If mvaridx > moPhones.Count Then
        mvarEOF = True
        Debug.Assert mvaridx > moPhones.Count
    End If
    FindCurrentIndex

End Sub

Private Sub IPhone_GotoPrevious()

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

Private Property Let IPhone_LocalNumber(ByVal RHS As String)

    moCurrent.Phone = RHS

End Property

Private Property Get IPhone_LocalNumber() As String

    IPhone_LocalNumber = moCurrent.Phone

End Property

Private Property Let IPhone_StoreTime(ByVal RHS As Date)

End Property

Private Property Get IPhone_StoreTime() As Date

End Property

Private Property Let IPhone_Type(ByVal RHS As String)

    If mvaridx = 0 Then
        IPhone_GotoLast
    End If

    If mvaridx > moPhones.Count Or moPhones.Count = 0 Then
        Dim gemsPh As GEMS_Phone
        Set gemsPh = New GEMS_Phone
        gemsPh.UpdateFlag = "N"
        
        moPhones.AddPhone gemsPh
        mvaridx = moPhones.Count
        FindCurrentIndex
    End If
    moCurrent.PhoneType = RHS
    'Type 0 means delete, 3M documentation states
    If Val(RHS) = 0 Then
        moCurrent.UpdateFlag = "D"
    End If
    
End Property

Private Property Get IPhone_Type() As String

    IPhone_Type = moCurrent.PhoneType
    
End Property
