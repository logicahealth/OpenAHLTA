VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'******************************************************
'WRogers
' Linked List nodes
'******************************************************
Private NodeNext    As Node
Private mvarData    As Variant
Private NodePrev    As Node

Public Event EndOfChain()

Friend Function GetPrevious() As Node

    If Not NodePrev Is Nothing Then
        Set GetPrevious = NodePrev
    Else
        RaiseEvent EndOfChain
    End If

End Function
'not valid to delete the first node in list
Public Function DeleteNode() As Node

    If NodeNext Is Nothing Then
        Set DeleteNode = Nothing
    Else
        Set DeleteNode = GetPrevious.GetNextNode
        GetNextNode.SetPreviousNode NodePrev
        GetPrevious.SetNextNode GetNextNode
        
    End If

End Function
Public Function GetNextNode() As Node

    If Not NodeNext Is Nothing Then
        Set GetNextNode = NodeNext
    Else
        RaiseEvent EndOfChain
    End If
        
End Function
Friend Sub SetPreviousNode(Prev As Node)

    Set NodePrev = Prev

End Sub
Public Function SetNextNode(ByRef Data As Variant) As Node
    
    If TypeOf Data Is Node Then
        Set NodeNext = Data
        Exit Function
    End If
    
    Set NodeNext = New Node
    
    NodeNext.SetData Data
    
    Set SetNextNode = NodeNext
    'to support double links
    NodeNext.SetPreviousNode Me
    
End Function
Public Sub SetData(Data As Variant)

    If IsObject(Data) Then
        Set mvarData = Data
    Else
        mvarData = Data
    End If

End Sub
Public Function GetData() As Variant

    If IsObject(mvarData) Then
        Set GetData = mvarData
    Else
        GetData = mvarData
    End If
    
End Function


Private Sub Class_Terminate()

    Set NodeNext = Nothing
    Set NodePrev = Nothing

End Sub
