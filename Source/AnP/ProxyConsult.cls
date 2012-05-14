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
Attribute VB_Name = "ProxyConsult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Handles all interaction with the Consult module.

Option Explicit

Private mConsult As Consult

Public Function DeleteConsult(ByVal DataID As String) As Boolean

  On Error GoTo ErrHandler

  InitConsultModule
      
  DeleteConsult = mConsult.DeleteConsultOrder(DataID)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyConsult.DeleteConsult", "A/P", vbExclamation

End Function

Private Sub InitConsultModule()
  If mConsult Is Nothing Then
    Set mConsult = gBB.ReadModule(NCID_CONSULT_TRACKING_APPLICATION)
  End If
End Sub

Private Function InitOps() As ConsultOrderOps
Dim Ops As ConsultOrderOps

  Set Ops = New ConsultOrderOps
    
  With Ops
    Set .Patient = gBB.Patient
    Set .Comm = gBB.Comm
    Set .Config = gBB.Config
    Set .Logon = gBB.Logon
    Set .User = gBB.User
  End With

  Set InitOps = Ops
  
End Function

Private Function LoadConsult(ByRef Consult As CHCSII_Consult.ConsultOrder) As iOrder
Dim AnPOrder As CHCSII_AnP.iOrder
Dim AnPConsult As CHCSII_AnP.OrderConsult
        
  On Error GoTo ErrHandler
        
  Set AnPOrder = New CHCSII_AnP.OrderConsult
  
  AnPOrder.Init Consult.ConsultId, Consult.ReceivingEntityDescription & ", " & Consult.PrimaryDxDescription

  Set AnPConsult = AnPOrder
  
  With AnPConsult
    'SCR 14684 2/20/02 Phil Crowder
    .ReferTo = Consult.ReceivingEntityDescription
    .Priority = Consult.UrgencyDescription
    .Clinic = Consult.RecvClinicName
    .OrderBy = gBB.Logon.UserName
    .OrderDate = Consult.OrderDate
    .Provider = Consult.OrderingProviderName
    .ProvisionalDiagnosis = Consult.PrimaryDxText     'SCR 32982  4/11/03  Phil Crowder
    .OrderID = Consult.OrderID
  End With
  
  Set LoadConsult = AnPOrder
        
  Exit Function
  
ErrHandler:
  Set LoadConsult = AnPOrder
  gShared.ShowVBError Err.Number, Err.Description, "ProxyConsult.LoadConsult", "A/P", vbExclamation
        
End Function

Public Function LoadConsults(ByVal EncounterID As String) As Collection
Dim Ops As ConsultOrderOps
Dim Consults As ConsultOrders
Dim Consult As CHCSII_Consult.ConsultOrder
Dim AnPOrders As Collection
Dim AnPOrder As iOrder
Dim i As Integer

  On Error GoTo ErrHandler
   
  Set AnPOrders = New Collection
  
  If Not gShared.IsAppMode(modeTheater) Then                'Theater doesn't use Consults
    InitConsultModule
    
    Set Ops = InitOps
    
    Set Consults = Ops.GetConsultsForEncounter(EncounterID)
    
    For i = 1 To Consults.Count
      Set Consult = Consults.Item(i)
      
      If Consult.OrderStatus <> Draft Then
        Set AnPOrder = LoadConsult(Consult)
        AnPOrders.Add AnPOrder, AnPOrder.Key
        
        Set AnPOrder = Nothing
      End If
    Next
  End If
  
  Set LoadConsults = AnPOrders
    
  Exit Function
   
ErrHandler:
  Set LoadConsults = AnPOrders
  gShared.ShowVBError Err.Number, Err.Description, "ProxyConsult.LoadConsults", "A/P", vbExclamation

End Function

Public Function ValidateAllConsultsSignable() As Boolean
Dim Ops As ConsultOrderOps

  On Error GoTo ErrHandler
  
  'SCR 46127 - Check for theater mode instead of localdb so we can perform the signable consult check in the CDR ITT mode
  If gShared.IsAppMode(modeTheater) Or gShared.IsAppMode(modeCHCSI_GUI) Then
    ValidateAllConsultsSignable = True
    Exit Function
  End If
  
  InitConsultModule
  
  Set Ops = InitOps
  
  ValidateAllConsultsSignable = Ops.IsSignable(gBB.CurrentEncounter.EncounterID)
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyConsult.ValidateAllConsultsSignable", "A/P", vbExclamation

End Function

