VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnPOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enumOrderType
  typAll
  typProc
  typConsult
  typLab
  typRad
  typMed
  typPI
  typInPat 'SCR 56524 KMonroe 11/03/04
End Enum

Private mKeyIndex As Long

Private mAllOrders As Collection

Private Sub Class_Initialize()
  Set mAllOrders = New Collection
End Sub

Public Function AddOrder(ByVal OrderType As enumOrderType) As iOrder
Dim Order As iOrder

  On Error GoTo ErrHandler
  
  Select Case OrderType
    Case typConsult
'      Set Order = New OrderConsult
      
    Case typLab
      Set Order = New OrderLab
    
    Case typMed
      Set Order = New OrderMed
      
    Case typPI
      Set Order = New OrderPatientInstruc
      
    Case typProc
      Set Order = New OrderProcedure
    
    Case typRad
      Set Order = New OrderRad
      
    Case typInPat  'SCR 56524 KMonroe 11/03/04
      Set Order = New OrderInPatient

  End Select
  
  AddOrderToCollection Order, OrderType
  
  Set AddOrder = Order
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.AddOrder", "A/P", vbExclamation
  
End Function

Private Sub AddOrderToCollection(ByRef Order As iOrder, ByRef OrderType As enumOrderType)
Dim Orders As Collection

  On Error GoTo ErrHandler
  
Retry:
  mKeyIndex = mKeyIndex + 1
  
  Set Orders = GetOrdersCollection(OrderType)
  Orders.Add Order, CStr(mKeyIndex)
  
  Order.Key = mKeyIndex
  
  Exit Sub
  
ErrHandler:
  If Err.Number <> 6 Then Resume Retry  '6 = Overflow
  
End Sub

Public Function Count(Optional ByVal OrderType As enumOrderType = typAll) As Long
Dim i As Long

  On Error GoTo ErrHandler
  
  If OrderType = typAll Then
    For i = 1 To mAllOrders.Count
      Count = Count + mAllOrders.Item(i).Count
    Next
  Else
    Count = mAllOrders.Item(CStr(OrderType)).Count
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.Count", "A/P", vbExclamation
  
End Function

Public Function Filter(Optional ByVal OrderType As enumOrderType = typAll, _
                        Optional ByVal FilterType As enumFilterType = 0) As Collection

  On Error GoTo ErrHandler

  If OrderType = typAll And FilterType = fltNone Then
    Set Filter = GetAllOrders

  ElseIf OrderType = typAll Then
    Set Filter = GetAllFilterOrders(FilterType)

  ElseIf FilterType = fltNone Then
    Set Filter = GetOrdersCollection(OrderType)

  Else
    Set Filter = GetSpecifiedOrdersFiltered(OrderType, FilterType)
  End If

  Exit Function

ErrHandler:
  Set Filter = New Collection
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.Filter", "A/P", vbExclamation

End Function

Private Function FilterOrders(ByVal FilterType As enumFilterType, ByRef Orders As Collection) As Collection
Dim i As Long
Dim Filtered As Collection

  On Error GoTo ErrHandler

  Set Filtered = New Collection

  Select Case FilterType
    Case fltActive
      For i = 1 To Orders.Count
        Select Case Orders.Item(i).State
          Case osNew, osSaved, osQueued, osModified
            Filtered.Add Orders.Item(i)
        End Select
      Next

    Case fltDirty
      For i = 1 To Orders.Count
        Select Case Orders.Item(i).State
          Case osNew, osDelete, osModified, osRemove
            Filtered.Add Orders.Item(i)
        End Select
      Next

    Case fltSaved
      For i = 1 To Orders.Count
        Select Case Orders.Item(i).State
          Case osSaved
            Filtered.Add Orders.Item(i)
        End Select
      Next

    Case fltQueued
      For i = 1 To Orders.Count
        Select Case Orders.Item(i).State
          Case osQueued
            Filtered.Add Orders.Item(i)
        End Select
      Next
  End Select

  Set FilterOrders = Filtered

  Exit Function

ErrHandler:
  Set FilterOrders = Filtered
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.FilterOrders", "A/P", vbExclamation

End Function

Private Function GetAllFilterOrders(ByVal FilterType As enumFilterType) As Collection
Dim Orders As Collection

  Set Orders = GetAllOrders
  Set GetAllFilterOrders = FilterOrders(FilterType, Orders)

End Function

Private Function GetAllOrders() As Collection
Dim AllOrders As Collection
Dim i As Long
Dim j As Long

  On Error GoTo ErrHandler

  Set AllOrders = New Collection

  For i = 1 To mAllOrders.Count
    For j = 1 To mAllOrders.Item(i).Count
      AllOrders.Add mAllOrders.Item(i).Item(j)
    Next
  Next

  Set GetAllOrders = AllOrders

  Exit Function

ErrHandler:
  Set GetAllOrders = AllOrders
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.GetAllOrders", "A/P", vbExclamation

End Function

Private Function GetOrdersCollection(ByRef OrderType As enumOrderType) As Collection
Dim Orders As Collection

  On Error Resume Next

  Set Orders = mAllOrders.Item(CStr(OrderType))

  If Orders Is Nothing Then
    Set Orders = New Collection
    mAllOrders.Add Orders, CStr(OrderType)
  End If

  Set GetOrdersCollection = Orders

End Function

Private Function GetSpecifiedOrdersFiltered(ByVal OrderType As enumOrderType, _
                                            ByVal FilterType As enumFilterType) As Collection
Dim Orders As Collection

  Set Orders = GetOrdersCollection(OrderType)
  Set GetSpecifiedOrdersFiltered = FilterOrders(FilterType, Orders)

End Function

Public Function Item(ByVal OrderType As enumOrderType, ByVal Key As Variant) As iOrder
  On Error Resume Next
  Set Item = mAllOrders.Item(CStr(OrderType)).Item(Key)
End Function

Public Function Load() As Boolean
Dim Success As Boolean
Dim EncounterID As String

  On Error GoTo ErrHandler
  
  EncounterID = CStr(gBB.CurrentEncounter.EncounterID)
  
  Success = LoadProcedures(EncounterID)
  Success = LoadLabs
  Success = LoadMeds
  Success = LoadRads
'  Success = LoadConsults(EncounterID)
  Success = LoadPIs
  Success = LoadInPats 'SCR 56524 KMonroe 11/03/04
  Load = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.Load", "A/P", vbExclamation
  
End Function

'Private Function LoadConsults(ByVal EncounterID As String) As Boolean
'Dim Proxy As ProxyConsult
'Dim Orders As Collection
'
'  On Error GoTo ErrHandler
'
'  Set Proxy = gBB.ConsultInterface
'  Set Orders = Proxy.LoadConsults(EncounterID)
'
'  mAllOrders.Add Orders, CStr(typConsult)
'
'  LoadConsults = True
'
'  Exit Function
'
'ErrHandler:
'  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadConsults", "A/P", vbExclamation
'
'End Function

Private Function LoadLabs() As Boolean
Dim Proxy As ProxyCDRPlus
Dim Orders As Collection

  On Error GoTo ErrHandler
  
  Set Proxy = gBB.CDRPlusInterface
  Set Orders = Proxy.LoadLabs
  
  mAllOrders.Add Orders, CStr(typLab)
  
  LoadLabs = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadLabs", "A/P", vbExclamation

End Function

Private Function LoadMeds() As Boolean
Dim Proxy As ProxyCDRPlus
Dim Orders As Collection

  On Error GoTo ErrHandler
  
  Set Proxy = gBB.CDRPlusInterface
  Set Orders = Proxy.LoadMeds
  
  mAllOrders.Add Orders, CStr(typMed)
  
  LoadMeds = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadMeds", "A/P", vbExclamation

End Function

Private Function LoadPIs() As Boolean
Dim CDRProxy As ProxyCDRPlus
Dim Orders As Collection

  On Error GoTo ErrHandler
  
  Set CDRProxy = gBB.CDRPlusInterface
  Set Orders = CDRProxy.LoadPIs
  
  mAllOrders.Add Orders, CStr(typPI)
  
  LoadPIs = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadPIs", "A/P", vbExclamation

End Function
'SCR 56524 KMonroe 11/03/04
Private Function LoadInPats() As Boolean
Dim CDRProxy As ProxyCDRPlus
Dim Orders As Collection

  On Error GoTo ErrHandler
  
  Set CDRProxy = gBB.CDRPlusInterface
  Set Orders = CDRProxy.LoadInPats
  
  mAllOrders.Add Orders, CStr(typInPat)
  
  LoadInPats = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadInPats", "A/P", vbExclamation

End Function

Private Function LoadProcedures(ByVal EncounterID As String) As Boolean
Dim Proxy As ProxyProblems
Dim Orders As Collection

  On Error GoTo ErrHandler
  
  Set Proxy = gBB.ProblemInterface
  
  Set Orders = Proxy.LoadProcedures(EncounterID)
  
  mAllOrders.Add Orders, CStr(typProc)
  
  LoadProcedures = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadProcedures", "A/P", vbExclamation
  
End Function

Private Function LoadRads() As Boolean
Dim Proxy As ProxyCDRPlus
Dim Orders As Collection

  On Error GoTo ErrHandler
  
  Set Proxy = gBB.CDRPlusInterface
  Set Orders = Proxy.LoadRads
  
  mAllOrders.Add Orders, CStr(typRad)
  
  LoadRads = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.LoadRads", "A/P", vbExclamation

End Function

Private Function OrdersRTF(ByRef Orders As Collection, ByVal SavedOnly As Boolean) As String
Dim Order As iOrder
Dim AddOrder As Boolean

  On Error GoTo ErrHandler
  
  For Each Order In Orders
    With Order
      If SavedOnly Then
        AddOrder = (.State = osSaved)
      Else
        AddOrder = (.State = osNew Or .State = osSaved Or .State = osModified)
      End If

      If AddOrder Then
        If Not .IsAssociated(SavedOnly) Then
          If InpatientEncounter = True Then
            OrdersRTF = OrdersRTF & BEGIN_ROW2 & INPAT_FONT_SIZE & .RTF & END_ROW
          Else
            OrdersRTF = OrdersRTF & BEGIN_ROW2 & .RTF & END_ROW
          End If
        End If
      End If
    End With
  Next
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.OrdersRTF", "A/P", vbExclamation
  Exit Function
  Resume
  
End Function

Private Sub RemoveOrders()
Dim i As Long
Dim j As Long

  On Error GoTo ErrHandler
  
  For i = 1 To mAllOrders.Count
    If mAllOrders.Item(i).Count > 0 Then
      For j = mAllOrders.Item(i).Count To 1 Step -1
        If mAllOrders.Item(i).Item(j).State = osRemove Then
          mAllOrders.Item(i).Remove j
        End If
      Next
    End If
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.RemoveOrders", "A/P", vbExclamation
  Resume Next
  
End Sub


Public Sub RemoveOrder(ByVal sOrderID As String)
Dim i As Long
Dim j As Long

  On Error GoTo ErrHandler
  
  For i = 1 To mAllOrders.Count
    If mAllOrders.Item(i).Count > 0 Then
      For j = mAllOrders.Item(i).Count To 1 Step -1
        If mAllOrders.Item(i).Item(j).DataID = sOrderID Then
          mAllOrders.Item(i).Item(j).Unassociate
          mAllOrders.Item(i).Remove j
        End If
      Next
    End If
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.RemoveOrder", "A/P", vbExclamation
  Resume Next
  
End Sub

Public Sub Reset()
  Set mAllOrders = Nothing
  Set mAllOrders = New Collection
End Sub

Public Function Save() As Boolean
Dim i As Long
Dim j As Long
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  Save = True
  
  For i = 1 To mAllOrders.Count
    For j = 1 To mAllOrders.Item(i).Count
      Success = mAllOrders.Item(i).Item(j).Save
      Save = Save And Success
    Next
  Next
  
  RemoveOrders
    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.Save", "A/P", vbExclamation
  
End Function

Public Function UnassociatedOrdersRTF(Optional ByVal SavedOnly As Boolean = False) As String

  On Error GoTo ErrHandler

  UnassociatedOrdersRTF = OrdersRTF(mAllOrders.Item(typProc), SavedOnly)
  UnassociatedOrdersRTF = UnassociatedOrdersRTF & OrdersRTF(mAllOrders.Item(CStr(typMed)), SavedOnly)
  UnassociatedOrdersRTF = UnassociatedOrdersRTF & OrdersRTF(mAllOrders.Item(CStr(typLab)), SavedOnly)
  UnassociatedOrdersRTF = UnassociatedOrdersRTF & OrdersRTF(mAllOrders.Item(CStr(typRad)), SavedOnly)
  'UnassociatedOrdersRTF = UnassociatedOrdersRTF & OrdersRTF(mAllOrders.Item(CStr(typConsult)), SavedOnly)
  UnassociatedOrdersRTF = UnassociatedOrdersRTF & OrdersRTF(mAllOrders.Item(CStr(typPI)), SavedOnly)
  UnassociatedOrdersRTF = UnassociatedOrdersRTF & OrdersRTF(mAllOrders.Item(CStr(typInPat)), SavedOnly) 'SCR 56524 KMonroe 11/03/04
 
  If Len(UnassociatedOrdersRTF) > 0 Then
    If InpatientEncounter = True Then
        UnassociatedOrdersRTF = END_LINE & INPAT_FONT_SIZE & " " & UNASSOC_HEADER & END_LINE & UnassociatedOrdersRTF
    Else
        UnassociatedOrdersRTF = END_LINE & " " & UNASSOC_HEADER & END_LINE & UnassociatedOrdersRTF
    End If
  End If

  Exit Function
  
ErrHandler:
    gShared.ShowVBError Err.Number, Err.Description, "AnPOrders.UnassociatedOrdersRTF", "A/P", vbExclamation
    Exit Function
    Resume
End Function
