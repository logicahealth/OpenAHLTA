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
Attribute VB_Name = "OrderAssociations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Manages all order association for a diagnosis.  Each order type can be associated to
'a diagnosis.  A collection of collections is used to store all of the Assoication
'objects.  The top level collection is for each type of order.  The 2nd level
'collection has all of the Assoication objects for a type of order.

Option Explicit

Private mAssociations As Collection

Private Sub Class_Initialize()
Dim Orders As Collection

  Set mAssociations = New Collection
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typProc)       'Procedures
  Set Orders = Nothing
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typConsult)    'Consults
  Set Orders = Nothing
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typLab)        'Labs
  Set Orders = Nothing
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typRad)        'Rads
  Set Orders = Nothing
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typMed)        'Meds
  Set Orders = Nothing
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typPI)         'Patient Instructions
  Set Orders = Nothing
  
  Set Orders = New Collection
  mAssociations.Add Orders, CStr(typInPat)         'Patient Instructions
  Set Orders = Nothing
  
End Sub

Public Function AssociateOrder(ByRef Assoc As OrderAssociation) As Boolean
  On Error Resume Next
  
  mAssociations.Item(CStr(Assoc.OrderType)).Add Assoc, Assoc.OrderKey
  
  If Err.Number > 0 Then
    mAssociations.Item(CStr(Assoc.OrderType)).Remove Assoc.OrderKey
    mAssociations.Item(CStr(Assoc.OrderType)).Add Assoc, Assoc.OrderKey
  End If
End Function

'SCR 28663  3/11/02  Phil Crowder  Changed RTF format for Labs
Private Function AssociationsRTF(ByRef Associations As Collection, ByRef Diagnosis As AnPDiagnosis, _
                                 ByVal Header As String, ByVal SavedOnly As Boolean) As String
Dim Association As OrderAssociation
Dim AddOrder As Boolean

  On Error GoTo ErrHandler
    
  For Each Association In Associations
    With Association
      If SavedOnly Then
        AddOrder = (.State = asAssociated Or .State = asUnassociate)
      Else
        AddOrder = (.State = asAssociated Or .State = asAssociate)
      End If
      
      If AddOrder And LenB(.RTF(Diagnosis)) > 0 Then
        If Len(AssociationsRTF) = 0 Then
          If Header <> LAB_HEADER Or Len(gShared.CmdLineSwitches.Item("GREEN")) = 0 Then
            AssociationsRTF = BEGIN_ROW & Header & NEW_CELL & .RTF(Diagnosis) & END_ROW
          Else
            AssociationsRTF = BEGIN_ROW & Header & NEW_CELL & .RTF(Diagnosis)
          End If
        Else
          If Header <> LAB_HEADER Or Len(gShared.CmdLineSwitches.Item("GREEN")) = 0 Then
            AssociationsRTF = AssociationsRTF & BEGIN_ROW & NEW_CELL & .RTF(Diagnosis) & END_ROW
          Else
            AssociationsRTF = AssociationsRTF & LAB_SEPERATOR & .RTF(Diagnosis)
          End If
        End If
      End If
    End With
  Next
    
  If Len(gShared.CmdLineSwitches.Item("GREEN")) > 0 Then
    If Header = LAB_HEADER And Len(AssociationsRTF) > 0 Then
      AssociationsRTF = AssociationsRTF & END_ROW
    End If
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociations.AssociationsRTF", "A/P", vbExclamation

End Function

Public Function Filter(Optional ByVal OrderType As enumOrderType = 0, _
                        Optional ByVal FilterType As enumFilterType = 0) As Collection
  On Error GoTo ErrHandler
  
  If OrderType = typAll And FilterType = fltNone Then
    Set Filter = GetAllAssociations
    
  ElseIf OrderType = typAll Then
    Set Filter = GetAllFilterAssociations(FilterType)
    
  ElseIf FilterType = fltNone Then
    Set Filter = GetAssoicationsCollection(OrderType)
  
  Else
    Set Filter = GetSpecifiedAssoicationsFiltered(OrderType, FilterType)
  End If
  
  Exit Function
  
ErrHandler:
  Set Filter = New Collection
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.Filter", "A/P", vbExclamation

End Function
 
Private Function FilterAssociations(ByVal FilterType As enumFilterType, ByRef Associations As Collection) As Collection
Dim i As Long
Dim Filtered As Collection
  
  On Error GoTo ErrHandler
  
  Set Filtered = New Collection
  
  Select Case FilterType
    Case fltActive
      For i = 1 To Associations.Count
        Select Case Associations.Item(i).State
          Case asAssociate, asAssociated, asQueued
            Filtered.Add Associations.Item(i)
        End Select
      Next
      
    Case fltDirty
      For i = 1 To Associations.Count
        Select Case Associations.Item(i).State
          Case asAssociate, asUnassociate
            Filtered.Add Associations.Item(i)
        End Select
      Next
    
    Case fltSaved
      For i = 1 To Associations.Count
        Select Case Associations.Item(i).State
          Case asAssociated
            Filtered.Add Associations.Item(i)
        End Select
      Next
    
    Case fltQueued
      For i = 1 To Associations.Count
        Select Case Associations.Item(i).State
          Case asQueued
            Filtered.Add Associations.Item(i)
        End Select
      Next
  End Select
  
  Set FilterAssociations = Filtered
  
  Exit Function
  
ErrHandler:
  Set FilterAssociations = Filtered
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociations.FilterOrders", "A/P", vbExclamation
  
End Function

Private Function GetAllFilterAssociations(ByVal FilterType As enumFilterType) As Collection
Dim Associations As Collection

  Set Associations = GetAllAssociations
  Set GetAllFilterAssociations = FilterAssociations(FilterType, Associations)
  
End Function

Private Function GetAllAssociations() As Collection
Dim Associations As Collection
Dim i As Long
Dim j As Long

  On Error GoTo ErrHandler

  Set Associations = New Collection
  
  For i = 1 To mAssociations.Count
    For j = 1 To mAssociations.Item(i).Count
      Associations.Add mAssociations.Item(i).Item(j)
    Next
  Next
  
  Set GetAllAssociations = Associations
  
  Exit Function
  
ErrHandler:
  Set GetAllAssociations = Associations
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociations.GetAllOrders", "A/P", vbExclamation
  
End Function

Private Function GetAssoicationsCollection(ByRef OrderType As enumOrderType) As Collection
Dim Associations As Collection

  On Error Resume Next
  
  Set Associations = mAssociations.Item(CStr(OrderType))
  
  If Associations Is Nothing Then
    Set Associations = New Collection
    mAssociations.Add Associations, CStr(OrderType)
  End If
  
  Set GetAssoicationsCollection = Associations
  
End Function

Private Function GetSpecifiedAssoicationsFiltered(ByVal OrderType As enumOrderType, _
                                            ByVal FilterType As enumFilterType) As Collection
Dim Associations As Collection

  Set Associations = GetAssoicationsCollection(OrderType)
  Set GetSpecifiedAssoicationsFiltered = FilterAssociations(FilterType, Associations)

End Function

Public Function Item(ByVal OrderType As enumOrderType, ByVal Key As Variant) As OrderAssociation
  On Error Resume Next
  Set Item = mAssociations.Item(CStr(OrderType)).Item(Key)
End Function

'JT BEGIN SCR 44681 - used to remove a corrupt association that references a missing order
'Can't use UnassociateOrder the order because the save would then delete it from the DB which would
'only make the order warning message available once.  We want it to show up each time the encounter
'is opened so that we don't lose the reference to the missing order
Public Sub RemoveAssocation(ByVal OrderType As enumOrderType, ByVal Key As Variant)
   On Error Resume Next
   mAssociations.Item(CStr(OrderType)).Remove Key
End Sub
'JT END SCR 44681

Public Sub RemoveAssociations()
Dim i As Long
Dim j As Long

  On Error GoTo ErrHandler
  
  For i = 1 To mAssociations.Count
    If mAssociations.Item(i).Count > 0 Then
      For j = mAssociations.Item(i).Count To 1
        If mAssociations.Item(i).Item(j).State = asUnassociated Then
          mAssociations.Item(i).Remove j
        End If
      Next
    End If
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociations.RemoveAssociations", "A/P", vbExclamation
  Resume Next

End Sub

Public Function RTF(ByRef Diagnosis As AnPDiagnosis, Optional ByVal SavedOnly As Boolean = False) As String

  On Error GoTo ErrHandler
    
  RTF = AssociationsRTF(mAssociations.Item(CStr(typProc)), Diagnosis, PROC_HEADER, SavedOnly)
  RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typMed)), Diagnosis, MED_HEADER, SavedOnly)
  
  'SCR 33819  3/25/03  Phil Crowder
  If Len(gShared.CmdLineSwitches.Item("GREEN")) > 0 Then
    RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typLab)), Diagnosis, LAB_HEADER, SavedOnly)
  Else
    RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typLab)), Diagnosis, LAB_HEADER_OLD, SavedOnly)
  End If
  
  'SCR 33819  3/25/03  Phil Crowder
  If Len(gShared.CmdLineSwitches.Item("GREEN")) > 0 Then
    RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typRad)), Diagnosis, RAD_HEADER, SavedOnly)
  Else
    RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typRad)), Diagnosis, RAD_HEADER_OLD, SavedOnly)
  End If
  
  RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typConsult)), Diagnosis, CONSULT_HEADER, SavedOnly)
  RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typPI)), Diagnosis, PI_HEADER, SavedOnly)
  RTF = RTF & AssociationsRTF(mAssociations.Item(CStr(typInPat)), Diagnosis, OTHER_ORDER_HEADER, SavedOnly)
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociations.RTF", "A/P", vbExclamation

End Function

Public Function Save(ByVal DiagnosisDataID As String) As Boolean
Dim i As Long
Dim j As Long
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  Save = True
  
  For i = 1 To mAssociations.Count
    For j = 1 To mAssociations.Item(i).Count
      Success = mAssociations.Item(i).Item(j).Save(DiagnosisDataID)
      
      Save = Save And Success
    Next
  Next
  
  RemoveAssociations
    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.Save", "A/P", vbExclamation

End Function
