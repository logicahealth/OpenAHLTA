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
Attribute VB_Name = "ParametersGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'A generic implementation of the iParameters interface.  Contains a collection of Parameter objects.

Option Explicit

Implements iObjectPersist
Implements iParameters

Private Const UNIQUE_IDENTIFIER As String = "ParametersGeneric"
Private Const OBJECT_VERSION As String = "1.0"

Private mParameters As Collection         'Collection of parameters

Private Sub Class_Initialize()
  Set mParameters = New Collection
End Sub

Private Sub Class_Terminate()
  Set mParameters = Nothing
End Sub

Private Function iObjectPersist_Copy() As Object
Dim NewMe As ParametersGeneric
Dim Persist As iObjectPersist
Dim Writer As iWriter
Dim Reader As iReader

  Set NewMe = New ParametersGeneric
  
  Set Writer = New WriterXML
  iObjectPersist_Persist Writer
  
  Set Reader = New ReaderXML
  Set Persist = NewMe

  Reader.DocumentVersion = Persist.Version
  Reader.Contents = Writer.Contents
    
  Persist.Restore Reader
  
  Set iObjectPersist_Copy = NewMe
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
Dim i As Long

  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION, mParameters.Count
    
    For i = 1 To mParameters.Count
      .WriteObject "Item_" & i, mParameters.Item(i)
    Next
  End With

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Obj As Parameter
Dim i As Long
Dim NumObjs As Long

  On Error GoTo ErrHandler
  
  Set mParameters = Nothing
  Set mParameters = New Collection
  
  With Reader
    NumObjs = .CollectionCount
    
    For i = 1 To NumObjs
      Set Obj = New Parameter
      .ReadObject "Item_" & i, Obj
      mParameters.Add Obj, Obj.Name
      
      Set Obj = Nothing
    Next
  End With

  iObjectPersist_Restore = True
  
  Exit Function
  
ErrHandler:

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As Parameter
Dim Param As Parameter

  On Error GoTo ErrHandler
  
  Set Param = New Parameter
  
  With Param
    .Name = Name
    .Value = Value
  End With
  
  mParameters.Add Param, Name
  
  Set iParameters_Add = Param
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "GenericParameters.iParameters_Add", Err.Description

End Function

Private Function iParameters_Count() As Long
  iParameters_Count = mParameters.Count
End Function

Private Function iParameters_Exists(ByVal Name As String) As Boolean
Dim Param As Parameter

  On Error Resume Next
  
  Set Param = mParameters.Item(Name)
  
  iParameters_Exists = (Err.Number = 0)
  
End Function

Private Function iParameters_Item(ByVal Key As Variant) As Parameter
  Set iParameters_Item = mParameters.Item(Key)
End Function

Private Sub iParameters_Remove(ByVal Key As Variant)
  On Error Resume Next
  mParameters.Remove Key
End Sub
