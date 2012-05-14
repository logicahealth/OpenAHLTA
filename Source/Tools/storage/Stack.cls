VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************
' WRogers
' 9/27/2001
'
' Stack data structure (a simple one)
'******************************************************
Private mvArray         As Variant
Private stackPointer    As Long
Private mlSize          As Long

Public Event StackOverFlow()
Public Event StackUnderFlow()

Public Sub Push(data As Variant)

   If IsEmpty(data) Then Exit Sub

    If stackPointer < mlSize Then
        stackPointer = stackPointer + 1
        If IsObject(data) Then
            Set mvArray(stackPointer) = data
        Else
            mvArray(stackPointer) = data
        End If
    Else
        RaiseEvent StackOverFlow
    End If

End Sub
Public Function Pop() As Variant

    If stackPointer < 0 Then
        RaiseEvent StackUnderFlow
    Else
        If IsObject(mvArray(stackPointer)) Then
            Set Pop = mvArray(stackPointer)
        Else
            Pop = mvArray(stackPointer)
        End If
        stackPointer = stackPointer - 1
    End If

End Function

Private Sub Class_Initialize()

    StackSize = 49 'default size

End Sub
'wipe out the stack if already used
Public Property Let StackSize(lSize As Long)

    mlSize = lSize
    stackPointer = -1
    ReDim mvArray(0 To mlSize)

End Property
Public Property Get StackSize() As Long

    StackSize = mlSize
    
End Property
