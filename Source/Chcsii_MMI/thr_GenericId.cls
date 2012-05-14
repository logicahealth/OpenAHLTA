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
Attribute VB_Name = "thr_GenericId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moGenericID As GEMS_Generic_IDs

Private moCurrent   As GEMS_Generic_Id
Private mvaridx     As Long
Private mvarBOF     As Boolean
Private mvarEOF     As Boolean

Implements IGenericId

Friend Property Set SetGenericID(genId As GEMS_Generic_IDs)

   Set moGenericID = genId
    mvarBOF = True
    mvaridx = 0
    
    If moGenericID.Count = 0 Then
        mvarEOF = True
    Else
        mvarEOF = False
    End If
    

End Property
Private Sub FindCurrentIndex()
    Dim lpos    As Long
    Dim v       As GEMS_Generic_Id

    lpos = 1
    For Each v In moGenericID

        lpos = lpos + 1
        If lpos > mvaridx Then
            Set moCurrent = v
            mvarBOF = False
            mvarEOF = False
            Exit For
        End If
    Next v

End Sub

Private Property Let IGenericId_AssigningAuthority(ByVal RHS As String)

    moCurrent.assigning_authority = RHS

End Property

Private Property Get IGenericId_AssigningAuthority() As String

    IGenericId_AssigningAuthority = moCurrent.assigning_authority

End Property

Private Property Get IGenericId_BOF() As Boolean
    IGenericId_BOF = mvarBOF
End Property

Private Property Get IGenericId_EOF() As Boolean
    IGenericId_EOF = mvarEOF
End Property

Private Sub IGenericId_GotoFirst()

    mvaridx = 1
    
    If moGenericID.Count = 0 Then
        mvarBOF = True
        mvarEOF = True
    ElseIf mvaridx <= moGenericID.Count Then
      mvarBOF = False
      mvarEOF = False
      FindCurrentIndex
    Else
      mvarEOF = True
      mvaridx = 0
    End If
    
    'Debug.Assert mvaridx

End Sub

Private Sub IGenericId_GotoLast()
    Dim v As Variant
    
    mvaridx = 0
    For Each v In moGenericID
        mvaridx = mvaridx + 1
    Next v
    
    FindCurrentIndex

End Sub

Private Sub IGenericId_GotoNext()

    mvaridx = mvaridx + 1
    If mvaridx > moGenericID.Count Then
        mvarEOF = True
        Debug.Assert mvaridx > moGenericID.Count
    End If
    FindCurrentIndex

End Sub

Private Sub IGenericId_GotoPrevious()

End Sub

Private Property Get IGenericId_StoreTime() As Date

    

End Property

Private Property Let IGenericId_StoreTime(ByVal RHS As Date)

End Property

Private Property Let IGenericId_Type(ByVal RHS As String)

    If mvaridx = 0 Then
        IGenericId_GotoLast
    End If

    If mvaridx > moGenericID.Count Or moGenericID.Count = 0 Then
        Dim gemsID As GEMS_Generic_Id
        Set gemsID = New GEMS_Generic_Id
        With gemsID
            .ID_TYPE_NCID = RHS
            .UpdateFlag = "N"
        End With
        
        moGenericID.AddGenericID gemsID
        mvaridx = moGenericID.Count
        FindCurrentIndex
    End If
    If Not IsNumeric(RHS) Then
        RHS = 0
    End If


    moCurrent.ID_TYPE_NCID = RHS
    'Type 0 means delete, 3M documentation states
    If Val(RHS) = 0 Then
        moCurrent.UpdateFlag = "D"
    End If
    
End Property

Private Property Get IGenericId_Type() As String

    IGenericId_Type = moCurrent.ID_TYPE_NCID

End Property

Private Property Let IGenericId_Value(ByVal RHS As String)

'scr 22757 Don't validate for numbers
'    If Not IsNumeric(RHS) Then
'        RHS = 0
'    End If

    moCurrent.ID_VALUE = RHS

End Property

Private Property Get IGenericId_Value() As String

    IGenericId_Value = moCurrent.ID_VALUE

End Property
