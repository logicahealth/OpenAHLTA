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
Attribute VB_Name = "thr_GenericData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moGenData   As GEMS_Generic_Datas
Private mvaridx     As Long
Private mvarBOF     As Boolean
Private mvarEOF     As Boolean
Private moCurrent   As GEMS_Generic_Data

Implements IGenericData

Friend Property Set GemsGenericData(gemsGenData As GEMS_Generic_Datas)

    Set moGenData = gemsGenData
    mvarBOF = True
    mvaridx = 0
    
    If moGenData.Count = 0 Then
        mvarEOF = True
    Else
        mvarEOF = False
    End If

End Property
    
Private Sub FindCurrentIndex()
    Dim lpos    As Long
    Dim v       As GEMS_Generic_Data

    lpos = 1
    For Each v In moGenData

        lpos = lpos + 1
        If lpos > mvaridx Then
            Set moCurrent = v
            Exit For
        End If
    Next v

End Sub

Private Property Get IGenericData_BOF() As Boolean

    IGenericData_BOF = mvarBOF

End Property

Private Property Let IGenericData_Container(ByVal RHS As Variant)

End Property

Private Property Get IGenericData_Container() As Variant

End Property

Private Property Get IGenericData_EOF() As Boolean

    IGenericData_EOF = mvarEOF

End Property

Private Property Let IGenericData_ErrorText(ByVal RHS As String)

End Property

Private Property Get IGenericData_ErrorText() As String

End Property

Private Sub IGenericData_GotoFirst()

    mvaridx = 1
    
    If moGenData.Count = 0 Then
        mvarBOF = True
        mvarEOF = True
    ElseIf mvaridx <= moGenData.Count Then
      mvarBOF = False
      mvarEOF = False
      FindCurrentIndex
    Else
      mvarEOF = True
      mvaridx = 0
    End If
    
End Sub

Private Sub IGenericData_GotoLast()
    Dim v As Variant
    
    mvaridx = 0
    For Each v In moGenData
        mvaridx = mvaridx + 1
    Next v
    
    FindCurrentIndex
    
End Sub

Private Sub IGenericData_GotoNext()

    If mvaridx > moGenData.Count Then
        mvarEOF = True
        
    End If
    mvaridx = mvaridx + 1
    FindCurrentIndex

End Sub

Private Sub IGenericData_GotoPrevious()

   If mvaridx = 1 Then
        mvarBOF = True
    ElseIf mvaridx = 0 Then
        Exit Sub
    End If
    
    mvaridx = mvaridx - 1
    
    Debug.Assert mvaridx >= 0
    
    FindCurrentIndex
    
End Sub

Private Property Let IGenericData_StoreTime(ByVal RHS As Date)

End Property

Private Property Get IGenericData_StoreTime() As Date

End Property

Private Property Let IGenericData_Type(ByVal RHS As String)
        
    If RHS = "" Then
        RHS = 0 'it's safe to overwrite my copy
    End If
    
    If mvaridx = 0 Then
        IGenericData_GotoLast
    End If
    
    If mvaridx > moGenData.Count Or moGenData.Count = 0 Then
        Dim gemData As GEMS_Generic_Data
        Set gemData = New GEMS_Generic_Data
        
        moGenData.AddGenericData gemData
        With gemData
            .Type_NCID = RHS
            .UpdateFlag = "N"
        End With
        mvaridx = moGenData.Count
        FindCurrentIndex
    End If
    
    moCurrent.Type_NCID = RHS
    'Type 0 means delete, 3M documentation states
    If Val(RHS) = 0 Then
        moGenData.Item(mvaridx).UpdateFlag = "D"
    End If
    
End Property

Private Property Get IGenericData_Type() As String

    IGenericData_Type = moCurrent.Type_NCID
    
End Property

Private Property Let IGenericData_ValueNcid(ByVal RHS As String)
    
    If RHS = "" Then
        RHS = 0 'it's safe to overwrite my copy
    End If
    
    moCurrent.Value_NCID = RHS

End Property

Private Property Get IGenericData_ValueNcid() As String

    IGenericData_ValueNcid = moCurrent.Value_NCID

End Property

Private Property Let IGenericData_ValueRsid(ByVal RHS As String)

    '?

End Property

Private Property Get IGenericData_ValueRsid() As String

    IGenericData_ValueRsid = 0

End Property

Private Property Let IGenericData_ValueText(ByVal RHS As String)

    moCurrent.Value_Text = RHS

End Property

Private Property Get IGenericData_ValueText() As String

    IGenericData_ValueText = moCurrent.Value_Text

End Property
