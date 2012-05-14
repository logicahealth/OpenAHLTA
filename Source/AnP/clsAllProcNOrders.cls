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
Attribute VB_Name = "clsAllProcNOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Name:    clsAllOrders.cls
'Author:        Phillip Crowder
'Created:       9/19/01
'Description:   The is a collection class for instances of clsItem as Orders.  It does
'               support functionality for all orders such as loading all orders
'               from problems and the CRD+,  saving all orders, etc.
'Revisions:
' 11/5/01       Moved all load functionality to clsOrdersFactory and moved all save
'               functionality to clsOrdersRepository.

Option Explicit

Private mcolOrders As Collection        'Collection of clsItem objects used as orders
Private mlngNextKey As Long             'Used to assign a unique key to each order

#If DEBUGON Then
  Private Const Module = "clsAllOrders.cls"
#End If

Private Sub Class_Initialize()
  Set mcolOrders = New Collection
  mlngNextKey = 1
End Sub

Private Sub Class_Terminate()
  Set mcolOrders = Nothing
End Sub

'This method is called to determine the number of active orders of a specified type
'in the collection of orders for the current encounter.
'Paramters:
' ItemType    Optional parameter to filter search
'Returns:
' Number of active orders
Public Property Get ActiveOrderCount(Optional ByVal ItemType As enumItemType = itUndefined) As Long
Dim objOrder As clsItem
Dim lngCount As Long

  On Error GoTo ErrHandler
  
  If ItemType = itUndefined Then
    'Count of all orders
    For Each objOrder In mcolOrders
      Select Case objOrder.ItemState
        Case isNew, isSaved, isQueued, isModified
          lngCount = lngCount + 1
      End Select
    Next
    
  Else
    For Each objOrder In mcolOrders
      'Count of a specific type of order
      If objOrder.ItemType = ItemType Then
        Select Case objOrder.ItemState
          Case isNew, isSaved, isQueued, isModified
            lngCount = lngCount + 1
        End Select
      End If
      
    Next
  End If
  
  ActiveOrderCount = lngCount
  
  Exit Property
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.ActiveOrderCount", "A/P", vbExclamation
  
End Property

'Checks to see if there are any new orders or orders that need to be deleted.
'Returns:
' True if there is one or more orders that need to saved or deleted and false if all
' orders have been saved or deleted.
Public Property Get IsDirty() As Boolean
Dim objOrder As clsItem
Dim bolDirty As Boolean

  On Error GoTo ErrHandler
   
  #If DEBUGON Then
     Dim DR  As DebugRec
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "IsDirty", "Begin", DR
  #End If

  For Each objOrder In mcolOrders
    Select Case objOrder.ItemState
      Case isNew, isDelete, isModified
        bolDirty = True
        Exit For
    End Select
  Next
  
  IsDirty = bolDirty
  
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "IsDirty", "End", DR
  #End If
  
  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllOrders.IsDirty", "A/P", vbExclamation
  
End Property

'Reference to whole collection of orders for the current encounter.
Public Property Get Orders() As Collection
  Set Orders = mcolOrders
End Property

'Adds a new order to the collection of orders and updates the unique key.
'Parameters:
' Order   New order object to add to the collection
Public Sub Add(ByRef Order As clsItem)

  On Error GoTo ErrHandler
  
  #If DEBUGON Then
     Dim DR  As DebugRec
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Add", "Begin", DR
  #End If

  Order.Key = mlngNextKey
  mlngNextKey = mlngNextKey + 1
  
  mcolOrders.Add Order, Order.Key
  
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Add", "End", DR
  #End If
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllOrders.Add", "A/P", vbExclamation
  
End Sub

'Used to change an existing queued order to a new order.  This is done after a queued
'order has been submitted to OE and OE returns a message to A/P that a new order
'has been created (see AnP.cls, Message function).
'Paramters:
' OrderKey    Key of Queued order
Public Sub AddQueuedOrder(ByVal OrderKey As String)
Dim objOrder As clsItem

  On Error GoTo ErrHandler
  
  Set objOrder = mcolOrders.Item(OrderKey)
  
  If objOrder.ItemState = isQueued Then
    objOrder.ItemState = isNew
  End If
    
  Set objOrder = Nothing
  
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.AddQueuedOrder", "A/P", vbExclamation
  
End Sub

'Remove deleted procedures and orders from collection of orders.  This should not be done
'until all orders and all diagnoses have been saved.
Public Sub CleanUpOrders()
Dim objOrder As clsItem
Dim strKey As String

  On Error GoTo ErrHandler
  
  #If DEBUGON Then
     Dim DR  As DebugRec
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CleanUpOrders", "Begin", DR
  #End If

  For Each objOrder In mcolOrders
  
    If objOrder.ItemState = isRemove Then
      strKey = objOrder.Key
      mcolOrders.Remove strKey
    End If
  
  Next

  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CleanUpOrders", "End", DR
  #End If

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.CleanUpOrders", "A/P", vbExclamation

End Sub
  
'Creates a new lab order from the OrderTablet in the OE module.
'Paramters:
' OEManager   OE module object used to get order values
' ItemState   state of order item when created
'Returns:
' Lab order object
Public Function CreateLabOrder(Optional ByRef OEManager As OE_SessionManager = Nothing, _
                                Optional ByVal ItemState As enumItemState = isNew) As clsItem
Dim objFactory As clsOrdersFactory

  Set objFactory = New clsOrdersFactory
  Set CreateLabOrder = objFactory.CreateLabOrder(OEManager, ItemState)
  Set objFactory = Nothing
  
End Function

'Creates a new med order from the OrderTablet in the OE module.
'Paramters:
' OEManager   OE module object used to get order values
' ItemState   state of order item when created
'Returns:
' Med order object
Public Function CreateMedOrder(Optional ByRef OEManager As OE_SessionManager = Nothing, _
                                Optional ByVal ItemState As enumItemState = isNew) As clsItem
Dim objFactory As clsOrdersFactory

  Set objFactory = New clsOrdersFactory
  Set CreateMedOrder = objFactory.CreateMedOrder(OEManager, ItemState)
  Set objFactory = Nothing

End Function

'Creates a new rad order from the OrderTablet in the OE module.
'Paramters:
' OEManager   OE module object used to get order values
' ItemState   state of order item when created
'Returns:
' Rad order object
Public Function CreateRadOrder(Optional ByRef OEManager As OE_SessionManager = Nothing, _
                                Optional ByVal ItemState As enumItemState = isNew) As clsItem
Dim objFactory As clsOrdersFactory

  Set objFactory = New clsOrdersFactory
  Set CreateRadOrder = objFactory.CreateRadOrder(OEManager, ItemState)
  Set objFactory = Nothing

End Function

'Used to get the Data Id for an order specified by key.
'Parameters:
' Key         Order Key to search for
'Returns:
' Data Id of order if found and null string if not found
Public Function FindDataId(ByVal Key As String) As String
  On Error Resume Next
  FindDataId = mcolOrders.Item(Key).PropValue(FLD_DATAID)
End Function

'Used to get the Order's key for a specified field and value.
'Paramters:
' IdField      search field
' IdValue      value to search for
' Active       If true then find orders that have an item state of new or saved and
'              if false look for queued orders
'Returns:
' Order Key of order if found and null string if not found
Public Function FindOrder(ByVal IdField As String, ByVal IdValue As String, Optional ByVal Active As Boolean = True) As String
Dim objOrder As clsItem

  On Error GoTo ErrHandler
  
  For Each objOrder In mcolOrders
    If objOrder.PropExists(IdField) Then
      If Active Then
        Select Case objOrder.ItemState
          Case isNew, isSaved, isModified
            'SCR 19767  12/17/01  Phil Crowder   Older templates saved orders in upper case
            If UCase$(Trim$(objOrder.PropValue(IdField))) = UCase$(Trim$(IdValue)) Then
              FindOrder = objOrder.Key
              Exit For
            End If
        End Select
        
      Else
        If objOrder.ItemState = isQueued Then
          'SCR 19767  12/17/01  Phil Crowder   Older templates saved orders in upper case
          If UCase$(Trim$(objOrder.PropValue(IdField))) = UCase$(Trim$(IdValue)) Then
            FindOrder = objOrder.Key
            Exit For
          End If
        End If
        
      End If
    End If
  Next
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.FindOrder", "A/P", vbExclamation
  
End Function

'Calls the order factory object which create a collection of orders for an encounter.
'Paramters:
' EncId         Used to filter orders for current encounter
' FacilityId    Used to filter orders for current encounter
Public Sub Load(ByVal EncId As String, ByVal FacilityID As String)
Dim objFactory As clsOrdersFactory

  On Error GoTo ErrHandler
  
  #If DEBUGON Then
     Dim DR  As DebugRec
     If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Load", "Begin", DR
  #End If
  
  Set objFactory = New clsOrdersFactory
  Set mcolOrders = objFactory.Load(EncId, FacilityID)
  Set objFactory = Nothing

  mlngNextKey = mcolOrders.Count + 1
  
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Load", "End", DR
  #End If
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.Load", "A/P", vbExclamation
  
End Sub

Public Function OrderExists(ByVal OrderKey As String) As Boolean
Dim objOrder As clsItem

  On Error Resume Next
  
  Set objOrder = mcolOrders.Item(OrderKey)
  OrderExists = Not (objOrder Is Nothing)
  
  Set objOrder = Nothing
  
End Function

Public Sub ResetOrders()
  On Error Resume Next
  
  mlngNextKey = 1
  
  Set mcolOrders = Nothing
  Set mcolOrders = New Collection
End Sub

'Calls the orders reposiotory object to save the collection of orders.
'Parameters:
' EncId         Saved to uniquely id orders for the encounter
' FacilityId    Saved to uniquely id orders for the encounter
'Returns:
' True if all orders where saved successfully and false on any errors.
Public Function Save(ByVal EncId As String, ByVal FacilityID As String) As Boolean
Dim objRepository As clsOrdersRepository
Dim bolOk As Boolean
Dim strMsg As String

  On Error GoTo ErrHandler

  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "Begin", DR
  #End If

  Set objRepository = New clsOrdersRepository
  bolOk = objRepository.Save(mcolOrders, EncId, FacilityID)
  
  If Not bolOk Then
    strMsg = objRepository.WarningMessage
  
    frmWarning.Display strMsg, True
    Set frmWarning = Nothing
  End If
  
  Save = bolOk
  
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "End", DR
  #End If
  
  Exit Function
  
ErrHandler:
  Save = False
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.Save", "A/P", vbExclamation
  
End Function

'This function is called when a Message is sent from OE module that an order has been renewed
'and/ or modified.  This method updates the Order's property values and sets its state
'to modified.
'Paramters:
' OEManager   OE Object used to get lab information
' Key         Key of order that is modified
Public Sub UpdateLabOrder(ByRef OEManager As OE_SessionManager, ByVal Key As String)
Dim objOrder As clsItem
      
  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UpdateLabOrder ", "Begin", DR
  #End If
      
  Set objOrder = mcolOrders.Item(Key)
  
  With OEManager.LabTablet
    If objOrder.ItemState = isSaved Then objOrder.ItemState = isModified
    objOrder.PropValue(FLD_ORDERID) = .OrderID
    objOrder.PropValue(FLD_COLLECT) = .CollectMethod
    objOrder.PropValue(FLD_COMMENT) = .Comment
    objOrder.PropValue(FLD_LAB_PRIORITY) = .LabCollectPriority
    objOrder.PropValue(FLD_P_PRIORITY) = .ProcessPriority
    objOrder.PropValue(FLD_SPEC_NAME) = .Specimen.Name
    objOrder.PropValue(FLD_SPEC_IEN) = .Specimen.IEN
    objOrder.PropValue(FLD_SPEC_TOP) = .Specimen.Topography
    objOrder.PropValue(FLD_SPEC_TOP_IEN) = .Specimen.TopographyIen
    objOrder.PropValue(FLD_SPEC_ANS_CD) = .Specimen.AnswerCode
    objOrder.PropValue(FLD_SPEC_ANS_TX) = .Specimen.AnswerText
    objOrder.PropValue(FLD_SPEC_CAP) = .Specimen.CapColorOrContainer
    objOrder.PropValue(FLD_TIME_ONE_CONT) = .Timing.OnexOrContinued
    objOrder.PropValue(FLD_TIME_ADMIN) = .Timing.AdminTime
    objOrder.PropValue(FLD_TIME_DUR) = .Timing.Duration
    objOrder.PropValue(FLD_TIME_FREQ) = .Timing.Frequency
    objOrder.PropValue(FLD_TIME_START) = .Timing.StartDate
    objOrder.PropValue(FLD_UNEXP_TIME) = .Timing.UnexpandedTime
    objOrder.PropValue(FLD_PROVIDER) = .OrderingProviderLname & ", " & .OrderingProviderFname
  End With

  Set objOrder = Nothing

  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UpdateLabOrder ", "End", DR
  #End If

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.UpdateLabOrder", "A/P", vbExclamation

End Sub

'This function is called when a Message is sent from OE module that an order has been renewed
'and/ or modified.  This method updates the Order's property values and sets its state
'to modified.
'Paramters:
' OEManager   OE Object used to get med information
' Key         Key of order that is modified
Public Sub UpdateMedOrder(ByRef OEManager As OE_SessionManager, ByVal Key As String)
Dim objOrder As clsItem

  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UpdateMedOrder ", "Begin", DR
  #End If

  Set objOrder = mcolOrders.Item(Key)
  
  With OEManager.RXTablet
    If objOrder.ItemState = isSaved Then objOrder.ItemState = isModified
    objOrder.PropValue(FLD_ORDERID) = .OrderID
    'Format function reports an overflow error on some dates if left in the yyyymmdd format
    'return by OE.  Why?
    objOrder.PropValue(FLD_START_DT) = Mid$(.OrderDate, 5, 2) & "/" & Right$(.OrderDate, 2) & "/" & Left$(.OrderDate, 4)
    objOrder.PropValue(FLD_CHILDCAP) = .ChildCap
    objOrder.PropValue(FLD_COMMENT) = .Comment
    objOrder.PropValue(FLD_SIG) = .SigCode
    objOrder.PropValue(FLD_QUANTITY) = .Quantity
    objOrder.PropValue(FLD_REFILLS) = .Refills
    objOrder.PropValue(FLD_PHAR_IEN) = .PharmacyIenSent
    objOrder.PropValue(FLD_PHAR_NAME) = .PharmacyNameSent
    objOrder.PropValue(FLD_PROVIDER) = .OrderingProviderLname & ", " & .OrderingProviderFname
  End With
  
  Set objOrder = Nothing
  
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UpdateMedOrder ", "End", DR
  #End If

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.UpdateMedOrder", "A/P", vbExclamation

End Sub

'This function is called when a Message is sent from OE module that an order has been renewed
'and/ or modified.  This method updates the Order's property values and sets its state
'to modified.
'Paramters:
' OEManager   OE Object used to get rad information
' Key         Key of order that is modified
Public Sub UpdateRadOrder(ByRef OEManager As OE_SessionManager, ByVal Key As String)
Dim objOrder As clsItem

  On Error GoTo ErrHandler
    
  #If DEBUGON Then
    Dim DR  As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UpdateRadOrder ", "Begin", DR
  #End If

  Set objOrder = mcolOrders.Item(Key)
  
  With OEManager.RadTablet
    If objOrder.ItemState = isSaved Then objOrder.ItemState = isModified
    objOrder.PropValue(FLD_ORDERID) = .OrderID
    objOrder.PropValue(FLD_COMMENT) = .Comment
    objOrder.PropValue(FLD_CLIN_IMPRESS) = .ClinicalImpressionText
    objOrder.PropValue(FLD_P_PRIORITY) = .ProcessPriority
    objOrder.PropValue(FLD_TRANS_MD) = .TransportMode
    objOrder.PropValue(FLD_WET_READ_WANT) = .WetReadingWanted
    objOrder.PropValue(FLD_TIME_ONE_CONT) = .Timing.OnexOrContinued
    objOrder.PropValue(FLD_TIME_ADMIN) = .Timing.AdminTime
    objOrder.PropValue(FLD_TIME_DUR) = .Timing.Duration
    objOrder.PropValue(FLD_TIME_FREQ) = .Timing.Frequency
    objOrder.PropValue(FLD_TIME_START) = .Timing.StartDate
    objOrder.PropValue(FLD_PROVIDER) = .OrderingProviderLname & ", " & .OrderingProviderFname
  End With

  Set objOrder = Nothing
  
  #If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "UpdateRadOrder ", "End", DR
  #End If

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.UpdateRadOrder", "A/P", vbExclamation

End Sub

Public Function UserHasPriv(ByVal Priv As CHCSII_PrivilegeType, Optional ByVal Key As String = vbNullString, Optional ByVal ItemType As enumItemType = itUndefined)
Dim User As Object
Dim Order As clsItem

  On Error Resume Next
  
  If Len(Key) > 0 Then
    Set Order = mcolOrders.Item(Key)
    Err.Clear
    If Order Is Nothing Then Exit Function
    
    ItemType = Order.ItemType
    Set Order = Nothing
  End If
  
  On Error GoTo ErrHandler
  
  
  Set User = gBB.User
  
  Select Case ItemType
  Case itProcedure
    UserHasPriv = User.HasPrivilegeEx(Priv_Procedure, Priv)
    
  Case itConsult
    UserHasPriv = User.HasPrivilegeEx(Priv_Orders_Consults, Priv)
    
  Case itRad
    UserHasPriv = User.HasPrivilegeEx(Priv_Orders_Radiology_Orders, Priv)
    
  Case itLab
    UserHasPriv = User.HasPrivilegeEx(Priv_Orders_Laboratory_Orders, Priv)
    
  Case itMed
    UserHasPriv = User.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv)
    
  Case itPatientInstr
    UserHasPriv = User.HasPrivilegeEx(Priv_Other_Therapies, Priv)
  
  Case Else
    UserHasPriv = True
  End Select
  
  Set User = Nothing
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "clsAllProcNOrders.UserHasDeletePriv", "A/P", vbCritical
  
End Function

