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
Attribute VB_Name = "WorkLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Application reference item in tree structure.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "WorkLoad"
Private Const OBJECT_VERSION As String = "1.0"

Private msAppt_Code As String        'Appointment Code
Private msWorkload_Type As String    'Provider Workload Type
Private msAppt_Duration As String    'Appointment Duration
Private msAppt_Description As String 'Appointment Description

Private Function iObjectPersist_Copy() As Object
Dim Item As WorkLoad
  On Error GoTo errHandler
  
  Set Item = New WorkLoad
  
  With Item
    .Appt_Code = msAppt_Code
    .Workload_Type = msWorkload_Type
    .Appt_Duration = msAppt_Duration
    .Appt_Description = msAppt_Description
  End With
  
  Set iObjectPersist_Copy = Item
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "DTOsWorkload.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
  On Error GoTo errHandler
  
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    .WriteProperty "Appt_Code", msAppt_Code
    .WriteProperty "Workload_Type", msWorkload_Type
    .WriteProperty "Appt_Duration", msAppt_Duration
    .WriteProperty "Appt_Description", msAppt_Description
  End With
  Exit Sub
    
errHandler:
    Err.Raise Err.Number, "DTOsWorkload.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo errHandler
  
  With Reader
    msAppt_Code = .ReadProperty("Appt_Code")
    msWorkload_Type = .ReadProperty("Workload_Type")
    msAppt_Duration = .ReadProperty("Appt_Duration")
    msAppt_Description = .ReadProperty("Appt_Description")
  End With

  iObjectPersist_Restore = Not Failed
  
  Exit Function
  
errHandler:
  Failed = True
  Resume Next

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property


Public Property Get Appt_Code() As String
    Appt_Code = msAppt_Code
End Property
Public Property Let Appt_Code(ByVal NewValue As String)
    msAppt_Code = NewValue
End Property
Public Property Get Appt_Duration() As String
    Appt_Duration = msAppt_Duration
End Property
Public Property Let Appt_Duration(ByVal NewValue As String)
    msAppt_Duration = NewValue
End Property
Public Property Get Workload_Type() As String
    Workload_Type = msWorkload_Type
End Property
Public Property Let Workload_Type(ByVal NewValue As String)
    msWorkload_Type = NewValue
End Property
Public Property Get Appt_Description() As String
    Appt_Description = msAppt_Description
End Property
Public Property Let Appt_Description(ByVal NewValue As String)
    msAppt_Description = NewValue
End Property


