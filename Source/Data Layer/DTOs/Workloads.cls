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
Attribute VB_Name = "Workloads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'NOTE: This class should NEVER be modified.

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "ProviderWorkloads"
Private Const OBJECT_VERSION As String = "1.0"

Private mCol As Collection

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Function iObjectPersist_Copy() As Object
Dim NewMe As Workloads
Dim Persist As iObjectPersist
Dim Writer As iWriter
Dim Reader As iReader

  On Error GoTo ErrHandler
  
  Set NewMe = New Workloads
  
  Set Writer = New WriterXML
  iObjectPersist_Persist Writer
  
  Set Reader = New ReaderXML
  Set Persist = NewMe
  
  Reader.DocumentVersion = Persist.Version
  Reader.Contents = Writer.Contents
  
  Persist.Restore Reader
  
  Set iObjectPersist_Copy = NewMe
  Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "DTOsWorkloads.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext
  
  
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
Dim i As Long
  On Error GoTo ErrHandler
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION, mCol.Count
    
    For i = 1 To mCol.Count
      .WriteObject "Item_" & i, mCol.Item(i)
    Next
  End With
  Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsWorkloads.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
  

End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
Dim Obj As Workload
Dim i As Long
Dim NumObjs As Long

  On Error GoTo ErrHandler
  
  With Reader
    NumObjs = .CollectionCount
    
    For i = 1 To NumObjs
      Set Obj = New Workload
      
      .ReadObject "Item_" & i, Obj
      mCol.Add Obj ', Obj.Appt_Code                'Note: If Collection has a Key, developer must add it here.
      
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

Public Sub Add(ByRef Item As Workload, Optional ByVal Key As String = vbNullString, Optional ByVal Before As Long)
    On Error Resume Next
    
    If Len(Key) > 0 Then
      If Before = 0 Then
          mCol.Add Item, Key
      Else
          If mCol.Count > 0 Then
              mCol.Add Item, Key, Before
          Else
              mCol.Add Item, Key
          End If
      End If
    Else
      If Before = 0 Then
          mCol.Add Item
      Else
          mCol.Add Item, , Before
      End If
    End If
    
End Sub

Public Function Count() As Long
  Count = mCol.Count
End Function

Public Function Item(ByVal Key As Variant) As Workload
Attribute Item.VB_UserMemId = 0
  Set Item = mCol.Item(Key)
End Function

Public Function Exists(ByVal Key As Variant) As Boolean
    Dim objWorkload As Workload
    
    On Error Resume Next
    Set objWorkload = mCol.Item(Key)
    
    If objWorkload Is Nothing Then
        Exists = False
    Else
        Exists = True
    End If
    
    Set objWorkload = Nothing
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(ByVal Key As Variant)
  On Error Resume Next
  mCol.Remove Key
End Sub

