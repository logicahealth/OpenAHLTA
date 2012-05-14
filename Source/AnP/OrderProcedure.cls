VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is sub-set of the Procedure in the problem module.  This class represents
'a Procedure specific to one Encounter (the problem module has all procedures for
'a patient).
'It has 2 main functions.  One, it contains the data needed to create the RTF for the
'Encounter Doc.  Two, its state can be changed without saving immediately (cache state
'changes).

Option Explicit

Implements iOrder

Private mOrder As iOrder        'Handles DataID, Key, and State code common to all Orders

Private mCPTCode As String
Private mCPTID As String
Private mNCID As String
Private msProviders As String

Private mComment As String
Private WithEvents mModifiers As ProcedureModifiers
Attribute mModifiers.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set mOrder = New iOrder
  
  mComment = Space$(1)
End Sub

Private Function iOrder_CreateAssociation() As OrderAssociation
Dim Assoc As OrderAssociation
  
  On Error GoTo ErrHandler

  Set Assoc = New OrderAssociation
  
  With Assoc
    .OrderKey = mOrder.Key
    .OrderType = typProc
    .OrderDataID = mOrder.DataID
    .State = asAssociate
  End With
  
  Set iOrder_CreateAssociation = Assoc
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderProcedure.iOrder_CreateAssociation", "A/P", vbExclamation

End Function

Private Property Get iOrder_DataID() As String
  iOrder_DataID = mOrder.DataID
End Property

Private Property Let iOrder_DataID(ByVal RHS As String)
  mOrder.DataID = RHS
End Property

Private Sub iOrder_Init(ByVal DataID As String, ByVal Name As String)
  mOrder.Init DataID, Name
End Sub

Private Function iOrder_IsAssociated(Optional ByVal SavedOnly As Boolean = False) As Boolean
  On Error Resume Next
  
  iOrder_IsAssociated = gBB.AllDiagnoses.IsOrderAssociated(typProc, mOrder.Key, SavedOnly)
End Function

Private Property Let iOrder_Key(ByVal RHS As String)
  mOrder.Key = RHS
End Property

Private Property Get iOrder_Key() As String
  iOrder_Key = mOrder.Key
End Property

Private Property Let iOrder_Name(ByVal RHS As String)
  mOrder.Name = RHS
End Property

Private Property Get iOrder_Name() As String
  iOrder_Name = mOrder.Name
End Property

'BR (29573) Modifer and Comment Format 1/17/03 Phil Crowder
'
'BR (28150) Display Modifers in Encounter Doc.  1/17/03  Phil Crowder
Private Property Get iOrder_RTF(Optional Diagnosis As AnPDiagnosis) As String

  iOrder_RTF = mOrder.Name
  
  If gShared.IsAppMode(modeCHCSI_GUI) Or gShared.IsAppMode(modeCDR) Then
          
      'SCR 38959 Revisited - a change was made to put the new format behind the /OPTIONS switch after this was coded
      'The ability to select the option is behind the switch but since the format of the RTF was changed,
      'we will need to call Display method instead of DisplayRTF when the swich is present so that we preserve the old format
      If Not (mModifiers Is Nothing) Then
            If Len(gShared.CmdLineSwitches("OPTION")) = 0 Then
               'Build this section the old way
               iOrder_RTF = iOrder_RTF & " " & mModifiers.Display
            Else
               'JT SCRs 38959 to 38960 - Let the modifier class determine where the code goes if the option was selected
               iOrder_RTF = iOrder_RTF & mModifiers.DisplayRTF(mCPTCode, gBB.EncounterInterface.IsIncludeRTFCodesSelected)
            End If
      ElseIf gBB.EncounterInterface.IsIncludeRTFCodesSelected Then
         'JT SCR 38959 - No modifier so just add the CPT4 Code to the RTF if option was selected
         iOrder_RTF = iOrder_RTF & " (" & mCPTCode & ")"
      End If
    
      If Len(Trim$(mComment)) > 0 Then
         'JT SCRs 38958 to 38960 - tag comments so they can be removed from sync warning
         iOrder_RTF = iOrder_RTF & RTF_COMMENT_BEGIN_TAG & " - " & RTFCodeFilter(mComment) & RTF_COMMENT_END_TAG      'JT SCR 40395 - convert the text to RTF coding
      End If
      
   End If
  
End Property

Private Property Let iOrder_State(ByVal RHS As enumOrderState)
  mOrder.State = RHS
End Property

Private Property Get iOrder_State() As enumOrderState
  iOrder_State = mOrder.State
End Property

Private Sub iOrder_RollbackState()
'Do nothing here
End Sub

Private Function iOrder_Save() As Boolean
Dim Required As Boolean
Dim Proxy As ProxyProblems
Dim DataID As String
Dim Success As Boolean
Dim oEnc As CHCSIIEncounterOps.Encounter
Dim oEncProv As EncProvider

  On Error GoTo ErrHandler

  Required = mOrder.State = osNew Or mOrder.State = osDelete Or mOrder.State = osModified
  
  If Not Required Then
    iOrder_Save = True
    Exit Function
  End If
  
  Set Proxy = gBB.ProblemInterface
  
  Select Case mOrder.State
    Case osNew
      DataID = Proxy.AddNewProcedure(mCPTCode)
      
      If Len(DataID) > 0 Then
        mOrder.DataID = DataID
        
        'Set the Order Data ID in all Associations of me
        gBB.AllDiagnoses.UpdateOrderDataID typProc, mOrder.Key, DataID
        
        'Set the Order Data ID in all provider associations of me
        If IsHIPAAEnabled Then
            Set oEnc = gBB.CurrentEncounter
            For Each oEncProv In oEnc.Providers
                oEncProv.UpdateProcedure DataID, mCPTID
            Next
        End If
        
        Success = True
      Else
        Success = False
      End If
    
      If ((Len(Trim$(mComment)) > 0 Or Modifiers.Codes.Count > 0 Or Modifiers.UnitsOfService > 1) Or gShared.IsAppMode(modeCHCSI_GUI)) And Success Then
        Success = Proxy.UpdateProcedure(Me)
      End If
      
    Case osDelete
      If Len(mOrder.DataID) > 0 Then
        Success = Proxy.DeleteProcedure(mOrder.DataID)
      Else
        Success = True      'Deleted but not saved to database yet
      End If
      
      'Unassociate the procedure from any providers
      If IsHIPAAEnabled Then
         Set oEnc = gBB.CurrentEncounter
         For Each oEncProv In oEnc.Providers
            oEncProv.UnAssociateProcedure DataID, mCPTID
         Next
      End If
      
    Case osModified
      Success = Proxy.UpdateProcedure(Me)
  End Select
  
  If Success Then mOrder.Save Else mOrder.RollbackState

  If mOrder.State = osRemove Then iOrder_Unassociate

  iOrder_Save = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderProcedure.iOrder_Save", "A/P", vbExclamation
  
End Function

'Set the State of All Associations of me to Unassociate
Private Sub iOrder_Unassociate()
Dim Diagnoses As AnPDiagnoses

  On Error GoTo ErrHandler

  Set Diagnoses = gBB.AllDiagnoses
  Diagnoses.UnassociateFromAll typProc, mOrder.Key

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderProcedure.iOrder_Unassociate", "A/P", vbExclamation

End Sub

Private Function iOrder_UserHasPriv(ByVal PrivType As CHCSII_ClientSecurity.CHCSII_PrivilegeType) As Boolean

  On Error GoTo ErrHandler

  iOrder_UserHasPriv = gBB.AnPChild(modProcedure).UserHasPriv(PrivType)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderProcedure.iOrder_UserHasPriv", "A/P", vbExclamation

End Function

Private Sub mModifiers_Modified()
  mOrder.State = osModified
End Sub

Public Property Let Comment(ByVal NewValue As String)
  If mComment <> Space$(1) Then mOrder.State = osModified
  
  mComment = Trim$(NewValue)
End Property

Public Property Get Comment() As String
  Comment = Trim$(mComment)
End Property

Public Property Let CPTCode(ByVal NewValue As String)
  mCPTCode = NewValue
End Property

Public Property Get CPTCode() As String
  CPTCode = mCPTCode
End Property

Public Property Let CPTID(ByVal NewValue As String)
  mCPTID = NewValue
End Property

Public Property Get CPTID() As String
  CPTID = mCPTID
End Property

Public Property Get Modifiers() As ProcedureModifiers
  If mModifiers Is Nothing Then
    Set mModifiers = New ProcedureModifiers
  End If
  
  Set Modifiers = mModifiers
End Property

Public Property Let NCID(ByVal NewValue As String)
  mNCID = NewValue
End Property

Public Property Get NCID() As String
  NCID = mNCID
End Property

Public Property Let SynonymName(ByVal NewValue As String)
  mOrder.Name = NewValue
End Property

Public Property Get SynonymName() As String
  SynonymName = mOrder.Name
End Property

Public Property Let Providers(ByVal sNewValue As String)
  msProviders = sNewValue
End Property

Public Property Get Providers() As String
  Providers = msProviders
End Property
