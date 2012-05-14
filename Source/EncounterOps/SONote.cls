VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SONote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This class is a data-holder only. See SO class for data i/o, etc.
Public Enum SONoteTypeEnum
    ListNoteType = 1
    FormNoteType = 2
End Enum

Public meType As SONoteTypeEnum
Public mnDataID As Long
Public mdDTS As Date
Public msOwnerNCID As String
Public msOwnerName As String
Public msTitle As String
Public msRTF As String
Public mbReadOnly As Boolean
Public mbNewCouplerData As Boolean
Public mbCouplerData As Boolean

Private mobjTextNote As TextNote  'SCR-48052

' Following members apply only to FormNoteType
Public msProfileNCID As String

' Following members apply only to ListNoteType
Public mcolListRecords As Collection

Friend Property Get DataId() As Long
    DataId = mnDataID
End Property
Friend Property Let DataId(lngDataID As Long)
    mnDataID = lngDataID
End Property

Public Sub ClearListRecords()
    Do While mcolListRecords.Count > 0
        mcolListRecords.Remove (1)
    Loop
End Sub
Public Property Get TextNote() As TextNote 'SCR-48052
    If mobjTextNote Is Nothing Then
        Set mobjTextNote = New TextNote
        mobjTextNote.msOwnerName = msOwnerName
        mobjTextNote.mdDTS = mdDTS
    End If
    Set TextNote = mobjTextNote
End Property
Public Property Set TextNote(oNote As TextNote)
    Set mobjTextNote = oNote
End Property

Friend Function Clone() As SONote
    Dim objNote As New SONote
    Dim objLTR As ListToolRecord
    With objNote
        .meType = meType
        .msOwnerName = msOwnerName
        .msOwnerNCID = msOwnerNCID
        .msProfileNCID = msProfileNCID
        .msRTF = msRTF
        .msTitle = msTitle
        .mdDTS = mdDTS
        Set .mcolListRecords = New Collection
        For Each objLTR In mcolListRecords
            .mcolListRecords.Add objLTR.Clone
        Next objLTR
    End With
    Set Clone = objNote
    Set objNote = Nothing
End Function

Friend Function RTF() As String

    Dim oShared As CWShared
    
    On Error Resume Next
    
    Set oShared = New CWShared
    
    'ensure proper date is used SCR 17620  WR 2/8/2002
    If mdDTS = 0 Then mdDTS = Now
    
    RTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
        & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
        & "{\colortbl;\red0\green0\blue0;}"
        
    'In PGUI Mode SO note may not exist 'SCR-48052
    If (Len(msRTF) And oShared.IsAppMode(modeCHCSI_GUI)) Or (Not oShared.IsAppMode(modeCHCSI_GUI)) Then
        If mbCouplerData And mbReadOnly Then
            RTF = RTF & "\pard\plain\f0\fs16\b\ul SO Note \b0 Cited by "
        Else
            RTF = RTF & "\pard\plain\f0\fs16\b\ul SO Note \b0 Written by "
        End If
            RTF = RTF & msOwnerName & " @ " & FormatDTG(mdDTS) & "\par\ul0\b "
        If msTitle <> "" Then
            RTF = RTF & msTitle & " \par "
        End If
        RTF = RTF & "\pard\plain \widctlpar\adjustright \cgrid" & msRTF & " \par "
    End If
    
    'In PGUI Append Text Note is it exists 'SCR-48052
    If oShared.IsAppMode(modeCHCSI_GUI) Then
        If Len(mobjTextNote.msRTF) Then
            If Len(msRTF) Then
                RTF = RTF & " \par "
            End If
            RTF = RTF & "\pard\plain\f0\fs16\b\ul SO Text Note \b0 Written by "
            RTF = RTF & msOwnerName & " @ " & FormatDTG(mdDTS) & "\par\ul0\b "
    
            RTF = RTF & "\pard\plain \widctlpar\adjustright \cgrid" & mobjTextNote.msRTF
        End If
    End If
    
    RTF = RTF & "}"
End Function
