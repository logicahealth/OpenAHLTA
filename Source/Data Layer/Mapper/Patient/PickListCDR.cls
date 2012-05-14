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
Attribute VB_Name = "PickListCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iPickList



Private Function iPickList_ComboxSelectStrings(ByVal sDomainNCID As String) As DTOs.PatientSelectStrings
    Dim objPickList As CONCEPTATLLib.PickListCtrl
    Dim oSelectString As DTOs.PatientSelectString
    Dim oSelectStrings As DTOs.PatientSelectStrings
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Set objPickList = New CONCEPTATLLib.PickListCtrl
    Set oSelectStrings = New DTOs.PatientSelectStrings
    
    '---set criteria for "loading" the pick-list
    objPickList.Context = NCID_3M_DEFAULT_CONTEXT
    objPickList.Domain = sDomainNCID
    objPickList.PrefOnlyFlag = True
    objPickList.Needed = True
    objPickList.Load
    
    For i = 1 To objPickList.ListItemCnt
        Set oSelectString = Nothing
        Set oSelectString = New DTOs.PatientSelectString
        
        oSelectString.ID = objPickList.ListItem(i).Concept
        oSelectString.Value = objPickList.ListItem(i).Representation
        oSelectStrings.Add oSelectString, oSelectString.ID
    Next i
    
    Set iPickList_ComboxSelectStrings = oSelectStrings
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "PickListCDR.iPickList_ComboxSelectStrings", Err.Description
CleanUp:
    Set oSelectString = Nothing
    Set oSelectStrings = Nothing
    Set objPickList = Nothing
End Function

Private Function iPickList_FMPSelectStrings(ByVal sNumDomainNCID As String, ByVal sTextDomainNCID As String) As DTOs.PatientSelectStrings

End Function

Private Function iPickList_PATCATSelectStrings(ByVal sNumDomainNCID As String) As DTOs.PatientSelectStrings

End Function

Private Function iPickList_ReturnRepFromNCID(ByVal sDomainNCID As String, ByVal sNCID As String) As String
    Dim oPicklist As New ConceptCtrl

    On Error GoTo ErrHandler
    
    iPickList_ReturnRepFromNCID = vbNullString
    
    If sNCID <> vbNullString Then
        oPicklist.UniqueId = sNCID
        iPickList_ReturnRepFromNCID = oPicklist.PrefRep(NCID_3M_DEFAULT_CONTEXT).Representation
    End If
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "PickListCDR.ReturnRepFromNCID", Err.Description
CleanUp:
    Set oPicklist = Nothing

End Function
