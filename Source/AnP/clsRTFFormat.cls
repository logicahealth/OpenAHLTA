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
Attribute VB_Name = "clsRTFFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Name:    clsRTFFormat.cls
'Author:        Phillip Crowder
'Created:       10/26/01
'Description:   This class is used to format a collection of diagnoses and orders into
'               a RTF string.
'Revisions:
' 4/11/02       Changed format of Orders.
' 4/13/02       Removed Sync Warning functions.  Created new class for them (clsSyncAnP.cls)

Option Explicit

#If DEBUGON Then
  Private Const Module = "clsRTFFormat.cls"
#End If

'Adds one consult order to the string of consult orders.
'Paramters:
' Order               Order object (clsItem) with the data to add to string of
'                     consult orders
' ConsultRTF          String of consult order to append the order information to
' LineHeader          Defines the left margin
' IncludeGroupHeader  True to include the 'Consult(s):' header with the 1st consult
'                     added to the string
Private Sub AppendConsult(ByRef Order As iOrder, ByRef ConsultRTF As String, _
                          ByRef LineHeader As String, ByVal IncludeGroupHeader As Boolean)
Dim strItem As String

  On Error GoTo ErrHandler

  If Len(ConsultRTF) = 0 And IncludeGroupHeader Then
    strItem = BEGIN_ROW & CONSULT_HEADER & NEW_CELL
  Else
    strItem = LineHeader
  End If

  strItem = strItem & Order.RTF
                  
  ConsultRTF = ConsultRTF & strItem & END_ROW
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation
  
End Sub

'Adds one lab order to the string of lab orders.
'Paramters:
' Order               Order object (clsItem) with the data to add to string of
'                     lab orders
' LabRTF              String of lab order to append the order information to
' LineHeader          Defines the left margin
' IncludeGroupHeader  True to include the 'Laboratory:' header with the 1st lab
'                     added to the string
Private Sub AppendLab(ByRef Order As iOrder, ByRef LabRTF As String, _
                      ByRef LineHeader As String, ByVal IncludeGroupHeader As Boolean)
Dim strItem As String

  On Error GoTo ErrHandler

  If Len(LabRTF) = 0 And IncludeGroupHeader Then
    strItem = BEGIN_ROW & LAB_HEADER & NEW_CELL
  Else
    strItem = LineHeader
  End If

  strItem = strItem & Order.RTF

  LabRTF = LabRTF & strItem & END_ROW
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation
  
End Sub

'Adds one med order to the string of med orders.
'Paramters:
' Order               Order object (clsItem) with the data to add to string of
'                     med orders
' MedRTF              String of med order to append the order information to
' LineHeader          Defines the left margin
' IncludeGroupHeader  True to include the 'Medication(s):' header with the 1st lab
'                     added to the string
Private Sub AppendMed(ByRef Order As iOrder, ByRef MedRTF As String, _
                      ByRef LineHeader As String, ByVal IncludeGroupHeader As Boolean)
Dim strItem As String

  On Error GoTo ErrHandler

  If Len(MedRTF) = 0 And IncludeGroupHeader Then
    strItem = BEGIN_ROW & MED_HEADER & NEW_CELL
  Else
    strItem = LineHeader
  End If
    
  strItem = strItem & Order.RTF

  MedRTF = MedRTF & strItem & END_ROW
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation
  
End Sub

'Adds one pi order to the string of pi orders.
'Paramters:
' Order               Order object (clsItem) with the data to add to string of
'                     pi orders
' PIRTF               String of pi order to append the order information to
' LineHeader          Defines the left margin
' IncludeGroupHeader  True to include the 'Patient Instruction(s):' header with the 1st lab
'                     added to the string
Private Sub AppendPI(ByRef Order As iOrder, ByRef PIRTF As String, _
                      ByRef LineHeader As String, ByVal IncludeGroupHeader As Boolean)
Dim strItem As String

  On Error GoTo ErrHandler

  If Len(PIRTF) = 0 And IncludeGroupHeader Then
     strItem = BEGIN_ROW & PI_HEADER & NEW_CELL
  Else
     strItem = LineHeader
  End If

  strItem = strItem & Order.RTF

  PIRTF = PIRTF & strItem & END_ROW

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation

End Sub

'Adds one procedure order to the string of procedure orders.
'Paramters:
' Order               Order object (clsItem) with the data to add to string of
'                     procedure orders
' ProcRTF             String of procedure order to append the order information to
' LineHeader          Defines the left margin
' IncludeGroupHeader  True to include the 'Procedure(s):' header with the 1st lab
'                     added to the string
Private Sub AppendProcedure(ByRef Order As iOrder, ByRef ProcRTF As String, _
                            ByRef LineHeader As String, ByVal IncludeGroupHeader As Boolean)
Dim strItem As String

  On Error GoTo ErrHandler
  
  If Len(ProcRTF) = 0 And IncludeGroupHeader Then
    strItem = BEGIN_ROW & PROC_HEADER & NEW_CELL
  Else
    strItem = LineHeader
  End If

  strItem = strItem & Order.RTF

  ProcRTF = ProcRTF & strItem & END_ROW

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation

End Sub

'Adds one rad order to the string of rad orders.
'Paramters:
' Order               Order object (clsItem) with the data to add to string of
'                     rad orders
' RadRTF              String of rad order to append the order information to
' LineHeader          Defines the left margin
' IncludeGroupHeader  True to include the 'Radiology:' header with the 1st lab
'                     added to the string
Private Sub AppendRad(ByRef Order As iOrder, ByRef RadRTF As String, _
                      ByRef LineHeader As String, ByVal IncludeGroupHeader As Boolean)
Dim strItem As String

  On Error GoTo ErrHandler
  
  If Len(RadRTF) = 0 And IncludeGroupHeader Then
    strItem = BEGIN_ROW & RAD_HEADER & NEW_CELL
  Else
    strItem = LineHeader
  End If

  strItem = strItem & Order.RTF

  RadRTF = RadRTF & strItem & END_ROW
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation
  
End Sub

'This method takes a diagnosis object and creates an RTF string of the diagnosis' data.
'Paramters:
' Diagnosis     Diagnosis object to use to create the RTF string
' IsPreview     True if the rft is to be previewed (ie. what the rtf will look like if all
'               current data is saved and false to generate the rtf for the encounter module.
'Returns:
' Diagnosis RTF string
Private Function CreateDiagnosis(ByRef Diagnosis As clsDiagnosis, ByVal IsPreview As Boolean) As String
Dim bolValid As Boolean
Dim strDiagnosis As String
Dim strComment As String

  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateDiagnosis", "Begin", DR
  #End If

  'Diagnosis is valid for display based on the IsPreview flag and the state of the
  'diagnosis.
  bolValid = IsValidDiagnosisState(Diagnosis, IsPreview)
    
  If bolValid Then
    strDiagnosis = strDiagnosis & DIAGNOSIS_HEADER & Diagnosis.Priority & _
                    DIAGNOSIS_PERIOD & Diagnosis.PropValue(FLD_SYNONYM)
    
    'Add extended name
    'SCR 17941  Phil Crowder 10/22/01
    'Added check for null value.
    If Len(Diagnosis.PropValue(FLD_EXTNAME)) > 0 Then
      strDiagnosis = strDiagnosis & DIAGNOSIS_EXT_NAME_HEADER & _
                      Diagnosis.PropValue(FLD_EXTNAME) & DIAGNOSIS_EXT_NAME_FOOTER
    Else
      strDiagnosis = strDiagnosis & DIAGNOSIS_FOOTER
    End If
    
    'Add comenets
    If Diagnosis.PropExists(FLD_COMMENT) Then
      strComment = RTFCodeFilter(Diagnosis.PropValue(FLD_COMMENT))
      strDiagnosis = strDiagnosis & DIAGNOSIS_COMMENT_HEADER & _
                      strComment & END_LINE
    End If
  End If
  
  CreateDiagnosis = strDiagnosis
  
  #If DEBUGON Then
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateDiagnosis", "End", DR
  #End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateDiagnosis", "A/P", vbExclamation
  
End Function

'Creates a string of all related items for one diagnosis.
'Paramters:
' Diagnosis     Diagnosis object with related items
' IsPreview     True if the rft is to be previewed (ie. what the rtf will look like if all
'               current data is saved and false to generate the rtf for the encounter module.
'Returns:
' String of related items
Private Function CreateRelatedItems(ByRef Diagnosis As clsDiagnosis, _
                                      ByVal IsPreview As Boolean) As String
Dim strConsults As String
Dim strLabs As String
Dim strMeds As String
Dim strProcedures As String
Dim strPIs As String
Dim strRads As String
Dim objItem As clsItem
Dim strKey As String
Dim Order As iOrder
Dim bolValid As Boolean

  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateRelatedItems", "Begin", DR
  #End If

  For Each objItem In Diagnosis.Items
    'An item is valid based on the IsPreview flag and the state of the item
    bolValid = IsValidItemState(objItem, IsPreview)
    
    If bolValid Then
      strKey = objItem.Key
      Set Order = gBB.Order(CovertItemTypeToOrderType(objItem.ItemType), strKey)
      
      Select Case objItem.ItemType
        Case itProcedure
          AppendProcedure Order, strProcedures, BEGIN_ROW & NEW_CELL, True
        Case itConsult
          AppendConsult Order, strConsults, BEGIN_ROW & NEW_CELL, True
        Case itRad
          AppendRad Order, strRads, BEGIN_ROW & NEW_CELL, True
        Case itLab
          AppendLab Order, strLabs, BEGIN_ROW & NEW_CELL, True
        Case itMed
          AppendMed Order, strMeds, BEGIN_ROW & NEW_CELL, True
        Case itPatientInstr
          AppendPI Order, strPIs, BEGIN_ROW & NEW_CELL, True
      End Select
    End If
    
  Next
    
  'The order of items was set by the customer
  CreateRelatedItems = strProcedures & strMeds & strLabs & strRads & strConsults & strPIs
  
  #If DEBUGON Then
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateRelatedItems", "End", DR
  #End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRelatedItems", "A/P", vbExclamation
  
End Function

'Creates the encounter RTF for all diagnoses and their related items.
'Parameters:
' AllDiagnoses  A collection of all diagnoses used to create RTF
' ApptDate      Used to determine if Order Date is added to the RTF for Orders
' IsPreview     True if the rft is to be previewed (ie. what the rtf will look like if all
'               current data is saved and false to generate the rtf for the encounter module.
'Returns:
' RTF string of all diagnoses, related items and unassociated orders
Public Function CreateRTF(ByRef AllDiagnoses As Collection, _
                          ByVal ApptDate As Date, Optional ByVal IsPreview As Boolean = False) As String
Dim strRTF As String
Dim strItems As String
Dim objDiagnosis As clsDiagnosis
Dim strDiagnosis As String

  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateRTF", "Begin", DR
  #End If
    
  strRTF = RTF_HEADER

  For Each objDiagnosis In AllDiagnoses
    'Add Diagnosis
    strDiagnosis = CreateDiagnosis(objDiagnosis, IsPreview)
    
    If Len(strDiagnosis) > 0 Then
      strRTF = strRTF & strDiagnosis
    
      'Add related items
      strItems = CreateRelatedItems(objDiagnosis, IsPreview)
      strRTF = strRTF & strItems
    End If
  Next
      
  'Add any unassociated items
  strItems = CreateUnassociatedItems(strRTF, IsPreview)
  
  CreateRTF = strRTF & strItems & RTF_FOOTER
  
  #If DEBUGON Then
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateRTF", "End", DR
  #End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateRTF", "A/P", vbExclamation
  Resume Next
  
End Function

'Creates a RTF string or orders that are not related to any diagnosis.
'Parameters:
' RTF           RTF string of all diagnoses and related
' IsPreview     True if the rft is to be previewed (ie. what the rtf will look like if all
'               current data is saved and false to generate the rtf for the encounter module.
'Returns:
' String of orders that are not assoicated to a diagnosis if any are found and a null
' string if all orders are assoicated to a least one diagnosis.
Private Function CreateUnassociatedItems(ByVal RTF As String, _
                                          ByVal IsPreview As Boolean) As String
Dim Orders As Collection
Dim Order As iOrder
Dim strItems As String
Dim bolValid As Boolean
Dim strProcedures As String
Dim strConsults As String
Dim strPIs As String
Dim strLabs As String
Dim strMeds As String
Dim strRads As String
Dim Consult As OrderConsult

  On Error GoTo ErrHandler
   
  #If DEBUGON Then
     Dim DR  As DebugRec
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateUnassociatedItems", "Begin", DR
  #End If

  Set Orders = gBB.AllOrders.Filter
  
  For Each Order In Orders
    'Order is valid based on the IsPreview flag and the state of the order
    bolValid = IsValidOrderState(Order, IsPreview)
        
    If bolValid Then
      
      With Order
        Select Case CovertOrderTypeToItemType(Order)
          Case itProcedure
            If InStr(1, RTF, .Name) = 0 Then
              AppendProcedure Order, strProcedures, BEGIN_ROW2, False
            End If
            
          Case itConsult
            Set Consult = Order
            
            If InStr(1, RTF, Consult.ReferTo) = 0 Then
              AppendConsult Order, strConsults, BEGIN_ROW2, False
            End If
            
          Case itPatientInstr
            If InStr(1, RTF, .Name) = 0 Then
              AppendPI Order, strPIs, BEGIN_ROW2, False
            End If
            
          Case itLab
            If InStr(1, RTF, .Name) = 0 Then
              AppendLab Order, strLabs, BEGIN_ROW2, False
            End If
            
          Case itMed
            If InStr(1, RTF, .Name) = 0 Then
              AppendMed Order, strMeds, BEGIN_ROW2, False
            End If
            
          Case itRad
            If InStr(1, RTF, .Name) = 0 Then
              AppendRad Order, strRads, BEGIN_ROW2, False
            End If
            
        End Select
      End With
      
    End If
  Next
  
  'The order of items displayed was specified by the customer
  strItems = strProcedures & strMeds & strLabs & strRads & strConsults & strPIs
    
  'If any unassociated orders are found then append them to the rtf string.
  If Len(strItems) > 0 Then
    strItems = END_LINE & " " & UNASSOC_HEADER & END_LINE & strItems
  End If

  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CreateUnassociatedItems", "End", DR
  #End If

  CreateUnassociatedItems = strItems
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsRTFFormat.CreateUnassociatedItems", "A/P", vbExclamation

End Function

'Used to determine if a diagnosis is valid and should be included in the RTF.
'Paramters:
' Diagnosis     Diagnosis object to check state of
' IsPreview     Mode that the RTF is being created in (True = Preview and False = Encounter)
'Returns:
' True if the diagnosis is valid and false if not
Private Function IsValidDiagnosisState(ByRef Diagnosis As clsDiagnosis, ByVal IsPreview As Boolean) As Boolean

  If Diagnosis Is Nothing Then Exit Function
  
  With Diagnosis
    If IsPreview Then
      IsValidDiagnosisState = (.DiagnosisState = dsDirty Or .DiagnosisState = dsNew Or .DiagnosisState = dsSaved)
    Else
      IsValidDiagnosisState = (.DiagnosisState = dsDirty Or .DiagnosisState = dsDelete Or .DiagnosisState = dsSaved)
    End If
  End With
  
End Function

'Used to determine if a related item is valid and should be included in the RTF.
'Paramters:
' Item          Item object to check state of
' IsPreview     Mode that the RTF is being created in (True = Preview and False = Encounter)
'Returns:
' True if the Item is valid and false if not
Private Function IsValidItemState(ByRef Item As clsItem, ByVal IsPreview As Boolean) As Boolean
Dim objOrder As iOrder

  On Error Resume Next
  
  If Item Is Nothing Then Exit Function
  
  Set objOrder = gBB.Order(CovertItemTypeToOrderType(Item.ItemType), Item.Key)
  
  If Not (objOrder Is Nothing) Then
    With Item
      If IsPreview Then
        IsValidItemState = (.ItemState = isAssociate Or .ItemState = isAssociated)
      Else
        IsValidItemState = (.ItemState = isAssociated Or .ItemState = isUnassociate)
      End If
    End With
  End If
  
End Function

'Used to determine if an order is valid and should be included in the RTF.
'Paramters:
' Order         Order object to check state of
' IsPreview     Mode that the RTF is being created in (True = Preview and False = Encounter)
'Returns:
' True if the Order is valid and false if not
Private Function IsValidOrderState(ByRef Order As iOrder, ByVal IsPreview As Boolean) As Boolean

  If Order Is Nothing Then Exit Function
  
  With Order
    If IsPreview Then
      IsValidOrderState = (.State = osNew Or .State = osSaved Or .State = osModified)
    Else
      IsValidOrderState = (.State = osSaved Or .State = osDelete)
    End If
  End With
  
End Function

'scr 21037 3/27/02 Phil Crowder
'Required for and text that the user can enter.  The \, and {} char are RTF special codes.
Private Function RTFCodeFilter(ByVal RTFString As String) As String

  RTFCodeFilter = Replace$(RTFString, "\", "\\")
  RTFCodeFilter = Replace$(RTFCodeFilter, "{", "\{")
  RTFCodeFilter = Replace$(RTFCodeFilter, "}", "\}")

End Function

