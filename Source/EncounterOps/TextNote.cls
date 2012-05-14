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
Attribute VB_Name = "TextNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This class is a data-holder only. See TextNotes class for data i/o, etc.

Public mnDataID As Long
Public mdDTS As Date
Public msOwnerNCID As String
Public msOwnerName As String
Public msCategory As String
Public msTitle As String
Public msRTF As String
Public mbIsComplete As Boolean
Private mbSensitive As Boolean
Private mbReadOnly As Boolean
Private meRelatedSection As EncounterSectionEnum ' SCR-38142
Private mnLockID As Long 'SCR#47028 Use lock id since user's ncid or note data id may be used to lock section
'''''''''''''''''''''''''''''''''''''''''''''''
' SCR #10622 Provide means of setting a       '
' Note as sensitve.                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Sensitive() As Boolean        '
    Sensitive = mbSensitive                      '''''''''''''''''''
End Property                                                        '
Public Property Let Sensitive(ByVal p_bSensitive As Boolean)        '
    mbSensitive = p_bSensitive                                     '
End Property                                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
' SCR 22188 Allow Note to be ReadOnly         '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ReadOnly() As Boolean
    ReadOnly = mbReadOnly
End Property
Public Property Let ReadOnly(ByVal bVal As Boolean)
    mbReadOnly = bVal
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
' Used to mark which section the note is related to
' Used by SO in PGUI mode SCR-38142
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get RelatedSection() As EncounterSectionEnum
    RelatedSection = meRelatedSection
End Property
Public Property Let RelatedSection(ByVal value As EncounterSectionEnum)
    meRelatedSection = value
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Property Get DataID() As Long
    DataID = mnDataID
End Property
Friend Property Let DataID(lngDataID As Long)
    mnDataID = lngDataID
End Property
Public Property Get LockID() As Long
    If mnLockID <> 0 Then
        LockID = mnLockID
    Else
        LockID = mnDataID
    End If
End Property
Public Property Let LockID(value As Long)
    mnLockID = value
End Property

Friend Function RTF() As String
    Dim sDraft As String
    On Error Resume Next
    RTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
        & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
        & "{\colortbl;\red0\green0\blue0;}"
    If Not mbIsComplete Then
        sDraft = " (Draft)"
    Else
        sDraft = ""
    End If
    RTF = RTF & "\pard\plain\f0\fs16\b\ul Note" & sDraft & " \b0 Written by " _
        & msOwnerName & " @ " & FormatDTG(mdDTS) & "\par\ul0\b "
    If msCategory <> "" Then
        RTF = RTF & msCategory & " \par "
    End If
    If msTitle <> "" Then
        RTF = RTF & msTitle & " \par "
    End If
    RTF = RTF & "\pard\plain \widctlpar\adjustright \cgrid" & msRTF & "}"
End Function

Friend Sub ReadFromDataString(ByRef objData As DATASTRATLLib.DataString)
    Dim objEvent As ClinicalEvent
    Dim objCommentGroup As CLINTYPEATLLib.CommentGroup
    Dim objCommentItem As CLINTYPEATLLib.CommentItem
    Dim sName As String
    Dim sValue As String
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    mnDataID = objData.DataID
    If Not objData.ModifyClinician Is Nothing Then
        msOwnerNCID = objData.ModifyClinician.ValueConcept.UniqueId
        msOwnerName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
    ElseIf Not objData.CreateClinician Is Nothing Then
        msOwnerNCID = objData.CreateClinician.ValueConcept.UniqueId
        msOwnerName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
    End If
    Set objEvent = objData.PatientData
    mdDTS = objEvent.Header.ObsDateTime.StartTime.value.DateTime
    If objEvent.ClinObs.Count > 0 Then
        If objEvent.ClinObs.Item(1).ObsMods.Count > 2 Then
            'RTF may be saved in sections SCR#42412
            'Checks to make sure RTF is stored in this ClinObs SCR#42412
            If StrComp(objEvent.ClinObs.Item(1).ObsMods(3).ObsValue.value.InfoType, "basicString", vbTextCompare) = 0 Then
                If StrComp(objEvent.ClinObs.Item(1).ObsMods(3).ObsValue.value.BasicString, "RTF", vbTextCompare) = 0 Then
                    msRTF = ExtractRTF2(objEvent, 1)
                End If
            End If
        Else
            msRTF = ExtractRTF(objEvent.ClinObs.Item(1))
        End If
    End If
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To objEvent.Header.Comments.Count
    'For Each objCommentGroup In objEvent.Header.Comments
        Set objCommentGroup = objEvent.Header.Comments(i)
        For j = 1 To objCommentGroup.Comments.Count
        'For Each objCommentItem In objCommentGroup.Comments
            Set objCommentItem = objCommentGroup.Comments(j)
            Call ParseCommentItem(objCommentItem, sName, sValue)
            Select Case sName
            Case "COMPLETE"
                If sValue = "Y" Then
                    mbIsComplete = True
                End If
            Case "CATEGORY"
                msCategory = sValue
            Case "TITLE"
                msTitle = sValue
            '''''''''''''''''''''''''''''''''''''''''''''''
            ' SCR #10622 Provide means of setting a       '
            ' Note as sensitve.                           '
            '''''''''''''''''''''''''''''''''''''''''''''''
            Case "SENSITIVE"                              '
                Sensitive = CBool(sValue)                 '
            '''''''''''''''''''''''''''''''''''''''''''''''
            Case "READONLY"
                mbReadOnly = CBool(sValue)
            End Select
        Next 'objCommentItem
    Next 'objCommentGroup
    Set objEvent = Nothing
End Sub

Friend Sub WriteToDataString(ByRef objData As DATASTRATLLib.DataString)
    Dim sComplete As String
    Dim sSensitive As String
    Dim sReadOnly As String
    'Call AddRTF(objData.PatientData, msRTF)
    Call AddRTF2(objData.PatientData, msRTF, 1) 'SCR#42412
    If mbIsComplete Then
        sComplete = "Y"
    Else
        sComplete = "N"
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' SCR #10622 Provide means of setting a       '
    ' Note as sensitve.                           '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSensitive = CStr(Sensitive)                                              '
    AddCommentValuePair objData.PatientData.Header, "SENSITIVE", sSensitive   '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sReadOnly = CStr(mbReadOnly)    'SCR 22188                                             '
    AddCommentValuePair objData.PatientData.Header, "READONLY", sReadOnly
    
    Call AddCommentValuePair(objData.PatientData.Header, "COMPLETE", sComplete)
    Call AddCommentValuePair(objData.PatientData.Header, "CATEGORY", msCategory)
    Call AddCommentValuePair(objData.PatientData.Header, "TITLE", msTitle)
End Sub
'Friend Sub GEMS_WriteToDataString(ByRef objData As ADODB.Recordset)
Friend Sub GEMS_WriteToDataString()

'GEMS   objdata is a rs of enc_rtf table
    Dim sComplete As String
    Dim vChunk As Variant
    Dim objSQL As ICHCSII_SQL.ISqlOpsEx
    Dim sReadOnly As String
    Dim oShared As CWShared
    Dim oSQLGen    As SQLGen
    Dim sSql       As String
    Dim objDAS           As ICHCSII_DAS
    Dim sbuf        As String
    Set oShared = New CWShared
    
    If mbIsComplete Then
        sComplete = "Y"
    Else
        sComplete = "N"
    End If
    
    If mbReadOnly Then 'Save Read Only flag SCR-786 ITT
        sReadOnly = "Y"
    Else
        sReadOnly = "N"
    End If
    
    'R. Foster SCR #48270;   Developer: Robert Foster 01/20/2004 05:57 PM
    'R. Foster Note: Function Changed:TextNote.cls - GEMS_WriteToDataString      Impact of Change:  Ensure that CHSSII_SQLOPS is set properly in ITT mode
    Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
    
    objSQL.Execute "select * from enc_rtfs where ENC_RTFSINDEX = " & mnDataID
    
    If Not objSQL.EOF Then
        objSQL("UserNCID") = msOwnerNCID
        objSQL("UserName") = msOwnerName
        objSQL("DTS") = mdDTS
        objSQL("complete") = sComplete
        objSQL("Category") = msCategory
        objSQL("Title") = msTitle
        objSQL("ReadOnly") = CStr(sReadOnly)  'SCR-22188
        'Compress RTF SCR-24683
        objSQL("OriginalSize") = Len(msRTF)
        
        'Compress only if RTF exists SCR-573 ITT
        If Len(msRTF) Then
            Dim objZlib As EncZLib
            Dim aBytes() As Byte
            Set objZlib = New EncZLib
            
            aBytes = StrConv(Replace(msRTF, Chr(0), ""), vbFromUnicode)
            objZlib.CompressData aBytes
            objSQL.value("Doc") = aBytes
        End If
        'objSQL("Doc") = CompressRTF(msRTF)
      'Call objSQL.AppendAsChunk("Doc", vChunk)
        objSQL.Update
    End If

    Set objSQL = Nothing
    
End Sub
Friend Sub GEMS_ReadFromRecordSet(ByRef objRS As ADODB.Recordset)
    Dim oConcept As New GEMS_ConceptCtrl
    Dim oShared As CWShared
    
    Set oShared = New CWShared

    On Error Resume Next
    If Not objRS.EOF Then
        mnDataID = Val("" & objRS("ENC_RTFSINDEX"))
        msOwnerNCID = "" & objRS("UserNCID")
        oConcept.UniqueId = msOwnerNCID
        msOwnerName = oConcept.PrefRep("2000").Representation
        mdDTS = "" & objRS("DTS")
        
        'msRTF = "" & objRS("Doc")
        If "" & objRS("complete") = "Y" Then
            mbIsComplete = True
        End If
        msCategory = "" & objRS("Category")
        msTitle = "" & objRS("TITLE")
        
        If "" & objRS("ReadOnly") = "Y" Then 'Load Read Only flag SCR-786 ITT
            mbReadOnly = True
        End If
        
        meRelatedSection = Val("" & objRS("RELATED_SECTION"))
        
        If oShared.IsAppMode(modeCHCSI_GUI) Then
            'msRTF = oShared.CHCSConnection.ReadChunk("enc_rtfs", "ENC_RTFSINDEX = " & mnDataID, "DOC")
            msRTF = ReadChunkCHCSConn("enc_rtfs", "ENC_RTFSINDEX = " & mnDataID, "DOC", False) 'SCR#43993
            mbReadOnly = True 'SCR-39840
        Else
            Dim objSQL As ICHCSII_SQL.ISqlOpsEx
            Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
            Dim vChunk As Variant
            objSQL.Execute "select doc from enc_rtfs where ENC_RTFSINDEX = " & mnDataID
            
            'Decompress if compressed SCR-24683
            If CVar(objRS("OriginalSize")) > 0 Then
                Dim objZlib As EncZLib
                Set objZlib = New EncZLib
                Dim aBytes() As Byte
                
                aBytes = objRS("Doc")
                Call objZlib.DecompressData(aBytes, objRS("OriginalSize"))
                msRTF = StrConv(aBytes, vbUnicode)
                'msRTF = DecompressRTF("" & objRS("DOC"), objRS("OriginalSize"))
            Else
                Call objSQL.GetAsChunk("Doc", vChunk)
                msRTF = vChunk
            End If
            Set objSQL = Nothing
        
        End If
    End If
End Sub

Friend Sub GEMS_WriteToRecordSet(ByRef objRS As ADODB.Recordset, Optional DeleteFlag As Boolean = False)

    If mnDataID = 0 Then
        objRS.AddNew
    Else
        Call objRS.Find("ENC_RTFSINDEX = " & mnDataID)
        If objRS.EOF Then
            objRS.AddNew
            objRS("ENC_RTFSINDEX") = mnDataID
        End If
    End If
    
    objRS("ownerncid") = msOwnerNCID
    objRS("DTS") = gobjCHCSIIConn.SQLDate(mdDTS)
    objRS("Doc") = msRTF
    If mbIsComplete Then
        objRS("complete") = "Y"
    Else
        objRS("complete") = "N"
    End If
    objRS("Category") = msCategory
    objRS("TITLE") = msTitle
    objRS.Update
    
    mnDataID = objRS("ENC_RTFSINDEX")
    
End Sub

