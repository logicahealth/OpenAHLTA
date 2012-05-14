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
Attribute VB_Name = "thr_Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moNames As Gems_Names
Private moCurrent As GEMS_Name
Private mvarBOF As Boolean
Private mvarEOF As Boolean
Private mvaridx As Long

Implements IName

Friend Property Set gemsName(gemsNames As Gems_Names)

    Set moNames = gemsNames
    mvarBOF = True
    mvaridx = 0

    If moNames.Count = 0 Then
        mvarEOF = True
    Else
        mvarEOF = False
    End If

End Property
Private Sub FindCurrentIndex()
    Dim lpos    As Long
    Dim v       As GEMS_Name

    lpos = 1
    For Each v In moNames

        lpos = lpos + 1
        If lpos > mvaridx Then
            Set moCurrent = v
            Exit For
        End If
    Next v

End Sub

Private Property Get IName_BOF() As Boolean

    IName_BOF = mvarBOF

End Property


Private Property Let IName_Degree(ByVal RHS As String)

    moCurrent.Degree = RHS

End Property

Private Property Get IName_Degree() As String

    IName_Degree = moCurrent.Degree

End Property

Private Property Get IName_EOF() As Boolean

    IName_EOF = mvarEOF

End Property

Private Property Let IName_First(ByVal RHS As String)

    moCurrent.First_name = RHS

End Property

Private Property Get IName_First() As String

    IName_First = moCurrent.First_name

End Property

Private Property Let IName_Full(ByVal RHS As String)

   '? moNames

End Property

Private Property Get IName_Full() As String

    IName_Full = moCurrent.Last_name & ", " & moCurrent.First_name & " " & moCurrent.Middle_name

End Property

'routine copied from GEMS_Addresses
Private Sub IName_GotoFirst()
    
    mvaridx = 1
    
    If moNames.Count = 0 Then
        mvarBOF = True
        mvarEOF = True
    ElseIf mvaridx <= moNames.Count Then
      mvarBOF = False
      mvarEOF = False
      FindCurrentIndex
    Else
      mvarEOF = True
      mvaridx = 0
    End If

End Sub

Private Sub IName_GotoLast()
    Dim v As Variant

    mvaridx = 0
    For Each v In moNames
        mvaridx = mvaridx + 1
    Next v

    FindCurrentIndex

End Sub

Private Sub IName_GotoNext()

    mvaridx = mvaridx + 1
    If mvaridx > moNames.Count Then
        mvarEOF = True
        Debug.Assert mvaridx > moNames.Count
    End If
    FindCurrentIndex

End Sub

Private Sub IName_GotoPrevious()

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

Private Property Let IName_Last(ByVal RHS As String)

    moCurrent.Last_name = RHS

End Property

Private Property Get IName_Last() As String

    IName_Last = moCurrent.Last_name

End Property

Private Property Let IName_Last2(ByVal RHS As String)

    moCurrent.Last_name2 = RHS

End Property

Private Property Get IName_Last2() As String

    IName_Last2 = moCurrent.Last_name2

End Property

Private Property Let IName_Middle(ByVal RHS As String)

    moCurrent.Middle_name = RHS

End Property

Private Property Get IName_Middle() As String

    IName_Middle = moCurrent.Middle_name

End Property

Private Property Let IName_Middle2(ByVal RHS As String)

    moCurrent.Middle_name2 = RHS

End Property

Private Property Get IName_Middle2() As String

    IName_Middle2 = moCurrent.Middle_name2

End Property

Private Property Let IName_StoreTime(ByVal RHS As Date)

End Property

Private Property Get IName_StoreTime() As Date

    

End Property

Private Property Let IName_Suffix(ByVal RHS As String)

    moCurrent.Suffix = RHS

End Property

Private Property Get IName_Suffix() As String

    moCurrent.Suffix = moCurrent.Suffix

End Property

Private Property Let IName_Title(ByVal RHS As String)

    moCurrent.Title = RHS

End Property

Private Property Get IName_Title() As String

    IName_Title = moCurrent.Title

End Property

Private Property Let IName_Type(ByVal RHS As String)
   
   If mvaridx = 0 Then
        IName_GotoLast
    End If

    If mvaridx > moNames.Count Or moNames.Count = 0 Then
        Dim gemsNm As GEMS_Name
        Set gemsNm = New GEMS_Name
        gemsNm.UpdateFlag = "N"
        
        moNames.AddName gemsNm
        mvaridx = moNames.Count
        FindCurrentIndex
    End If
    moCurrent.Name_type_cid = RHS
    'Type 0 means delete, 3M documentation states
    If Val(RHS) = 0 Then
        moCurrent.UpdateFlag = "D"
    End If
    
End Property

Private Property Get IName_Type() As String

    IName_Type = moCurrent.Name_type_cid

End Property
