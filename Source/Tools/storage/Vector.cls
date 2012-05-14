VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'******************************************************
' WRogers
' 4/15/2002
'
' vector - dynamically sized one dimensional array
'******************************************************
Private mnHeadEnd   As storage.Node
Private mlPosition   As Long
Private mnCurrent   As storage.Node
Private mlCount      As Long

Public Event PastEnd()
'returns a postion Index
Public Function Add(Data As Variant) As Long

    If mlPosition = -1 Then
        mnHeadEnd.SetData Data
        Set mnCurrent = mnHeadEnd
    Else
        'find the end
        Last
        Set mnCurrent = mnCurrent.SetNextNode(Data)
    End If
        
    mlPosition = mlPosition + 1
    mlCount = mlCount + 1

    Add = mlPosition

End Function
Function Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    
    If Index < mlCount And Index > -1 Then
        If Index = mlPosition Then

        ElseIf Index > mlPosition Then
            Do While Index > mlPosition
                mlPosition = mlPosition + 1
                Set mnCurrent = mnCurrent.GetNextNode
            Loop
        Else 'must be Index < mlPosition
            Do While Index < mlPosition And mlPosition > 0
                mlPosition = mlPosition - 1
                Set mnCurrent = mnCurrent.GetPrevious
            Loop

        End If
        If IsObject(mnCurrent.GetData) Then
            Set Item = mnCurrent.GetData
        Else
            Item = mnCurrent.GetData
        End If

    Else
        RaiseEvent PastEnd
    End If

End Function
'overwrite data at current position, return previous data, if any
Public Function Replace(Data As Variant) As Variant

    If mlPosition > -1 Then
        If IsObject(mnCurrent.GetData) Then
            Set Replace = mnCurrent.GetData
        Else
            Replace = mnCurrent.GetData
        End If

        mnCurrent.SetData Data
    End If

End Function
'return previous data
Public Function Delete(Index As Long) As Variant

    If Index < mlCount And Index > -1 Then
        If Index = mlPosition Then
        ElseIf Index > mlPosition Then
            Do While Index > mlPosition
                mlPosition = mlPosition + 1
                Set mnCurrent = mnCurrent.GetNextNode
            Loop
        Else 'must be Index < mlPosition
            Do While Index < mlPosition
                mlPosition = mlPosition - 1
                Set mnCurrent = mnCurrent.GetPrevious
            Loop
        End If
        
        If IsObject(mnCurrent.GetData) Then
            Set Delete = mnCurrent.GetData
        Else
            Delete = mnCurrent.GetData
        End If
        
        mlCount = mlCount - 1
        
        mnCurrent.DeleteNode
        
        First
            
    Else
        RaiseEvent PastEnd
    End If

End Function
Public Function Length() As Long

    Length = mlCount

End Function
Public Function PreviousData() As Variant

    If mlPosition > 0 Then
        mlPosition = mlPosition - 1
        Set mnCurrent = mnCurrent.GetPrevious
        If IsObject(mnCurrent.GetData) Then
            Set PreviousData = mnCurrent.GetData
        Else
            PreviousData = mnCurrent.GetData
        End If

    Else
        RaiseEvent PastEnd
    End If

End Function
'returns data from next mlPosition. Next is a better name, but Next is a keyword.
Public Function NextData() As Variant

    If mlPosition < mlCount - 1 Then
        mlPosition = mlPosition + 1
        Set mnCurrent = mnCurrent.GetNextNode
        If IsObject(mnCurrent.GetData) Then
            Set NextData = mnCurrent.GetData
        Else
            NextData = mnCurrent.GetData
        End If

    Else
        RaiseEvent PastEnd
    End If

End Function
Public Function First() As Variant

    If mlCount > 0 Then
        If IsObject(mnCurrent.GetData) Then
            Set First = mnCurrent.GetData
        Else
            First = mnCurrent.GetData
        End If

'        First = mnHeadEnd.GetData
        Set mnCurrent = mnHeadEnd
        mlPosition = 0
    End If

End Function
Public Function Last() As Variant

    If mlCount > 0 Then
        Do While mlPosition < mlCount - 1
            mlPosition = mlPosition + 1
            Set mnCurrent = mnCurrent.GetNextNode
        Loop
        If IsObject(mnCurrent.GetData) Then
            Set Last = mnCurrent.GetData
        Else
            Last = mnCurrent.GetData
        End If
    End If

End Function
Private Sub Class_Initialize()

    Set mnHeadEnd = New Node
    mlPosition = -1 'before first element

End Sub

Private Sub Class_Terminate()

    Set mnHeadEnd = Nothing
    Set mnCurrent = Nothing

End Sub
