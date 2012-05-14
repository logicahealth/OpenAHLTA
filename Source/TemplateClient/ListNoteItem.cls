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
Attribute VB_Name = "ListNoteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private miFlag As Integer
Private msSnoID As String
Private msPrefix As String
Private msNote As String
Private msResult As String
Public Property Get Flag() As Integer
    
    On Error GoTo ErrHandler
    Flag = miFlag
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Flag", Err.Description

End Property

Public Property Let Flag(ByVal RHV As Integer)
    
    On Error GoTo ErrHandler
    miFlag = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Flag", Err.Description

End Property

Public Property Get SnoID() As String
    
    On Error GoTo ErrHandler
    SnoID = msSnoID
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.SnoID", Err.Description

End Property

Public Property Let SnoID(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msSnoID = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.SnoID", Err.Description

End Property

Public Property Get Prefix() As String
    
    On Error GoTo ErrHandler
    Prefix = msPrefix
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Prefix", Err.Description

End Property

Public Property Let Prefix(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msPrefix = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Prefix", Err.Description

End Property

Public Property Get Note() As String
    
    On Error GoTo ErrHandler
    Note = msNote
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Note", Err.Description

End Property

Public Property Let Note(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msNote = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Note", Err.Description

End Property


Public Property Get Result() As String
    
    On Error GoTo ErrHandler
    Result = msResult
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Result", Err.Description

End Property

Public Property Let Result(ByVal RHV As String)
    
    On Error GoTo ErrHandler
    msResult = RHV
    Exit Property

ErrHandler:
    Err.Raise Err.Number, "CHCSII_TemplateClient.ListNoteItem.Result", Err.Description

End Property
