VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProxyOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Handles all interaction with the OE and GEM_OE modules.

Option Explicit

Private mCDRSessionMgr As CHCSII_OE.OE_SessionManager
Private mTheaterSessionMgr As GEMS_OE.OE_SessionManager

'OE method only.
Public Property Get Tablet() As Object
  
  On Error GoTo ErrHandler
  
  InitCDROE

  Select Case mCDRSessionMgr.OrderTypeCode
    Case hlRX_ORDERS_CODE
      Set Tablet = mCDRSessionMgr.RXTablet
      
    Case hlLAB_ORDERS_CODE
      Set Tablet = mCDRSessionMgr.LabTablet
      
    Case hlRAD_ORDERS_CODE
      Set Tablet = mCDRSessionMgr.RadTablet
  End Select
  
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.Tablet", "A/P", vbExclamation
  
End Property

'OE method only.
Public Sub AddDeletedMedNote()

  On Error GoTo ErrHandler

  InitCDROE
  
  If mCDRSessionMgr.OrderTypeCode = hlRX_ORDERS_CODE Then
    gBB.EncounterInterface.AddEncounterNote "Discontinued " & _
                                            mCDRSessionMgr.RXTablet.OrderName & " - " & _
                                            mCDRSessionMgr.RXTablet.SigCode
  End If
  
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.AddDeletedMedNote", "A/P", vbExclamation
  
End Sub

Private Sub AddEncounterNote(ByVal OrderKey As String)
Dim Med As OrderMed

  On Error GoTo ErrHandler
  
  If Len(OrderKey) = 0 Then Exit Sub
  
  Set Med = gBB.Order(typMed, OrderKey)
  
  gBB.EncounterInterface.AddEncounterNote "Discontinued " & _
                                          Med.OrderName & " - " & Med.SigCode
  
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.AddEncounterNote", "A/P", vbExclamation

End Sub

'OE method Only.
'Without check with the OE module it is not possible to determine the type of Order
'that was created by OE.
Public Function CreateOrder() As iOrder
Dim Order As iOrder

  On Error GoTo ErrHandler
        
  InitCDROE
  
  Select Case mCDRSessionMgr.OrderTypeCode
    Case hlRX_ORDERS_CODE
      Set Order = gBB.AllOrders.AddOrder(typMed)
      UpdateMedOrder Order
      
    Case hlLAB_ORDERS_CODE
      Set Order = gBB.AllOrders.AddOrder(typLab)
      UpdateLabOrder Order
      
    Case hlRAD_ORDERS_CODE
      Set Order = gBB.AllOrders.AddOrder(typRad)
      UpdateRadOrder Order
  End Select
  
  Set CreateOrder = Order

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.CreateOrder", "A/P", vbExclamation

End Function

Public Function DeleteOrder(ByVal AnPOrderType As enumOrderType, ByVal OrderID As String, _
                            Optional ByVal OrderKey As String = vbNullString) As Boolean
Dim OrderType As hlORDERTYPECODE
Dim Success As Boolean

  On Error GoTo ErrHandler

  OrderType = TranslateOrderType(AnPOrderType)
  
  Success = DeleteTheaterOrder(OrderType, OrderID)
    
  If Success And OrderType = hlRX_ORDERS_CODE Then
    'SCR 21424 Phil Crowder 4/24/02
    AddEncounterNote OrderKey
  End If
    
  DeleteOrder = Success
  
  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.DeleteOrder", "A/P", vbExclamation
    
End Function

'GEM_OE method only.
Private Function DeleteTheaterOrder(ByVal OrderType As hlORDERTYPECODE, ByVal OrderID As String) As Boolean
Dim ErrCode As hlERRORS
  
  On Error GoTo ErrHandler
  
  InitTheaterOE
      
  ErrCode = mTheaterSessionMgr.PerformAction(OrderType, hlACTION_DELETE, OrderID)

  DeleteTheaterOrder = (ErrCode = hlERROR_NONE)

  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.DeleteTheaterOrder", "A/P", vbExclamation

End Function

'SetDeleteOE is a flag in the Lab Order object that indicates that the order
'was deleted in OE already.  Does a search by Order ID and Order Name (for Queued
'Orders).
Private Function FindLab(ByVal SetDeletedOE As Boolean) As iOrder
Dim Orders As Collection
Dim i As Long
Dim Lab As OrderLab
Dim Order As iOrder

  On Error GoTo ErrHandler

  Set Orders = gBB.AllOrders.Filter(typLab, fltActive)
  
  For i = 1 To Orders.Count
    Set Lab = Orders.Item(i)
    
    If Lab.OrderID = mCDRSessionMgr.LabTablet.OrderID Then
      Lab.DeletedInOE = SetDeletedOE
      Set FindLab = Lab
      Exit For
    
    ElseIf UCase$(Lab.OrderName) = UCase$(mCDRSessionMgr.LabTablet.OrderName) Then
      Set Order = Lab
      
      If Order.State = osQueued Then
        Lab.DeletedInOE = SetDeletedOE
        Set FindLab = Lab
        Exit For
      End If
    End If
  Next
  
  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.FindLab", "A/P", vbExclamation
  
End Function

'SetDeleteOE is a flag in the Order object that indicates that the order
'was deleted in OE already.
Public Function FindMatchingOrder(Optional ByVal SetDeletedOE As Boolean = False) As iOrder

  On Error GoTo ErrHandler

  InitCDROE
  
  Select Case mCDRSessionMgr.OrderTypeCode
    Case hlLAB_ORDERS_CODE
      Set FindMatchingOrder = FindLab(SetDeletedOE)
      
    Case hlRX_ORDERS_CODE
      Set FindMatchingOrder = FindMed(SetDeletedOE)
    
    Case hlRAD_ORDERS_CODE
      Set FindMatchingOrder = FindRad(SetDeletedOE)
  End Select
  
  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.FindMatchingOrder", "A/P", vbExclamation
  
End Function

'SetDeleteOE is a flag in the Med Order object that indicates that the order
'was deleted in OE already.  Does a search by Order ID and Order Name (for Queued
'Orders).
Private Function FindMed(ByVal SetDeletedOE As Boolean) As iOrder
Dim Orders As Collection
Dim i As Long
Dim Med As OrderMed
Dim Order As iOrder

  On Error GoTo ErrHandler

  Set Orders = gBB.AllOrders.Filter(typMed, fltActive)
  
  For i = 1 To Orders.Count
    Set Med = Orders.Item(i)
    
    If Med.OrderID = mCDRSessionMgr.RXTablet.OrderID Then
      Med.DeletedInOE = SetDeletedOE
      Set FindMed = Med
      Exit For
    
    ElseIf UCase$(Med.OrderName) = UCase$(mCDRSessionMgr.RXTablet.OrderName) Then
      Set Order = Med
      
      If Order.State = osQueued Then
        Med.DeletedInOE = SetDeletedOE
        Set FindMed = Med
        Exit For
      End If
    End If
  Next
  
  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.FindMed", "A/P", vbExclamation
  
End Function

'SetDeleteOE is a flag in the Rad Order object that indicates that the order
'was deleted in OE already.  Does a search by Order ID and Order Name (for Queued
'Orders).
Private Function FindRad(ByVal SetDeletedOE As Boolean) As iOrder
Dim Orders As Collection
Dim i As Long
Dim Rad As OrderRad
Dim Order As iOrder

  On Error GoTo ErrHandler

  Set Orders = gBB.AllOrders.Filter(typRad, fltActive)
  
  For i = 1 To Orders.Count
    Set Rad = Orders.Item(i)
    
    If Rad.OrderID = mCDRSessionMgr.RadTablet.OrderID Then
      Rad.DeletedInOE = SetDeletedOE
      Set FindRad = Rad
      Exit For
    
    ElseIf UCase$(Rad.OrderName) = UCase$(mCDRSessionMgr.RadTablet.OrderName) Then
      Set Order = Rad
      
      If Order.State = osQueued Then
        Rad.DeletedInOE = SetDeletedOE
        Set FindRad = Rad
        Exit For
      End If
    End If
  Next
  
  Exit Function
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.FindRad", "A/P", vbExclamation
  
End Function

'OE method only
Private Sub InitCDROE()
Dim CDR_Get_Manager As CHCSII_OE.OE_GetSessionManager
Dim CHCS2Interface As ICHCSII_CoreAPI2

  On Error GoTo ErrHandler

  If Not (mCDRSessionMgr Is Nothing) Then Exit Sub
  
  Set CDR_Get_Manager = New CHCSII_OE.OE_GetSessionManager
  Set mCDRSessionMgr = CDR_Get_Manager.GetSessionManager
  
  Set CHCS2Interface = mCDRSessionMgr
  
  With gBB
    CHCS2Interface.Initialize2 0, .Comm, .Logon, .Config, .Patient, .User
  End With
        
  Exit Sub
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.InitCDROE", "A/P", vbExclamation
        
End Sub

'GEM_OE method only.
Private Sub InitTheaterOE()
Dim Theater_Get_Manager As GEMS_OE.OE_GetSessionManager

  On Error GoTo ErrHandler

  If Not (mTheaterSessionMgr Is Nothing) Then Exit Sub

  Set Theater_Get_Manager = New GEMS_OE.OE_GetSessionManager
  Set mTheaterSessionMgr = Theater_Get_Manager.GetSessionManager

  Exit Sub
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.InitTheaterOE", "A/P", vbExclamation

End Sub

Private Function TranslateOrderType(ByVal AnPOrderType As enumOrderType) As hlORDERTYPECODE
  
  Select Case AnPOrderType
    Case typLab
      TranslateOrderType = hlLAB_ORDERS_CODE
      
    Case typMed
      TranslateOrderType = hlRX_ORDERS_CODE
      
    Case typRad
      TranslateOrderType = hlRAD_ORDERS_CODE
  End Select
  
End Function

Private Sub UpdateLabOrder(ByRef Order As OrderLab)
      
  On Error GoTo ErrHandler
  
  With mCDRSessionMgr.LabTablet
    Order.OrderID = .OrderID
    Order.OrderName = .OrderName
    Order.OrderIEN = .OrderIEN
    Order.CollectionMethod = .CollectMethod
    Order.Comment = .Comment
    Order.LabCollectionPriority = .LabCollectPriority
    Order.ProcessPriority = .ProcessPriority
    Order.SpecimenName = .Specimen.Name
    Order.SpecimenIEN = .Specimen.IEN
    Order.SpecimenTogography = .Specimen.Topography
    Order.SpecimenTogographyIEN = .Specimen.TopographyIen
    Order.SpecimenAnswerCode = .Specimen.AnswerCode
    Order.SpecimenAnswerText = .Specimen.AnswerText
    Order.SpecimenContainer = .Specimen.CapColorOrContainer
    Order.TimingOneOrContinued = .Timing.OnexOrContinued
    Order.TimingAdminTime = .Timing.AdminTime
    Order.TimingDuration = .Timing.Duration
    Order.TimingFrequence = .Timing.Frequency
    Order.TimingStartDate = .Timing.StartDate
    Order.UnexpandedTime = .Timing.UnexpandedTime
    Order.Provider = .OrderingProviderLname & ", " & .OrderingProviderFname
    Order.OrderedBy = gBB.Logon.UserName
    Order.TimeOrdered = Now
  End With

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.UpdateLabOrder", "A/P", vbExclamation

End Sub

Private Sub UpdateMedOrder(ByRef Order As OrderMed)
    
  On Error GoTo ErrHandler
  
  With mCDRSessionMgr.RXTablet
    Order.OrderID = .OrderID
    Order.OrderIEN = .OrderIEN
    Order.OrderName = .OrderName
    'Format function reports an overflow error on some dates if left in the yyyymmdd format
    'return by OE.  Why?
    Order.StartDate = Mid$(.OrderDate, 5, 2) & "/" & Right$(.OrderDate, 2) & "/" & Left$(.OrderDate, 4)
    Order.ChildCap = .ChildCap
    Order.Comment = .Comment
    Order.SigCode = .SigCode
    Order.Quantity = .Quantity
    Order.Refills = .Refills
    Order.PharmacySentIEN = .PharmacyIenSent
    Order.PharmacySent = .PharmacyNameSent
    Order.Provider = .OrderingProviderLname & ", " & .OrderingProviderFname
    Order.OrderedBy = gBB.Logon.UserName
    Order.TimeOrdered = Now
  End With

  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.UpdateMedOrder", "A/P", vbExclamation

End Sub

Public Sub UpdateOrder(ByRef Order As iOrder)
  
  On Error GoTo ErrHandler
        
  InitCDROE
  
  Select Case mCDRSessionMgr.OrderTypeCode
    Case hlRX_ORDERS_CODE
      UpdateMedOrder Order
      
    Case hlLAB_ORDERS_CODE
      UpdateLabOrder Order
      
    Case hlRAD_ORDERS_CODE
      UpdateRadOrder Order
  End Select

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.UpdateOrder", "A/P", vbExclamation

End Sub

Private Sub UpdateRadOrder(ByRef Order As OrderRad)
    
  On Error GoTo ErrHandler
  
  With mCDRSessionMgr.RadTablet
    Order.OrderID = .OrderID
    Order.OrderName = .OrderName
    Order.OrderIEN = .OrderIEN
    Order.Comment = .Comment
    Order.ClinicalImpression = .ClinicalImpressionText
    Order.ProcessPriority = .ProcessPriority
    Order.TransportionMode = .TransportMode
    Order.WetReadingWanted = .WetReadingWanted
    Order.TimingOneOrContinued = .Timing.OnexOrContinued
    Order.TimingAdminTime = .Timing.AdminTime
    Order.TimingDuration = .Timing.Duration
    Order.TimingFrequence = .Timing.Frequency
    Order.TimingStartDate = .Timing.StartDate
    Order.Provider = .OrderingProviderLname & ", " & .OrderingProviderFname
    Order.OrderedBy = gBB.Logon.UserName
    Order.TimeOrdered = Now
  End With

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyOE.UpdateRadOrder", "A/P", vbExclamation

End Sub
