VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mintSBEReqFlag As Integer
Private mstrComments As String
Private mstrRTFComments As String
Private mintSBE As Integer
Private mblnDisplayRequiredSBE As Boolean
Private mlngDentEncId As Long
Private mlngPatUnitId As Long
Private mblnDirty As Boolean

Implements IntrfcCHCSII_Dental.IDENTAL_SBE

Public Sub Set_SBE_Info(ByVal intSBE As Integer, ByVal strComments As String)
    mstrComments = strComments
    mintSBE = intSBE
    
    'SCR 48874 and 48875 SF IDENTAL_SBE_GetUserInput accounts for a value of 4 but this sub did not.
    If mintSBE = 0 Or mintSBE = 4 Then
       mintSBEReqFlag = 0
    Else
       mintSBEReqFlag = -1
    End If
End Sub
Public Sub Set_SBEReq_Flag(ByVal intSBEReqFlag As Integer)
    mintSBEReqFlag = intSBEReqFlag
End Sub

Private Sub SaveSBE()
On Error GoTo ErrHandler

'    If mDParent Is Nothing Then
'      DentalPatient_Search = False
'      Err.Raise ErrBase_frmSBE, VBModule, "DentalPatient_Search - Dental Parent Not Set"
'      Exit Sub  'Go Bye Bye
'    End If
      
    Dim strSQL As String
    Dim objDAS As ICHCSII_DAS
    Dim objRS As ADODB.Recordset
            
    'Find if Patient exist in DENTAL_PATIENT table
    Set objDAS = GetDasObj()
    strSQL = "Update DENTAL_ENCOUNTER Set SBE=" & mintSBE & ", SBE_COMMENTS = '" & _
          RQ(mstrComments) & "' Where DENTAL_ENCOUNTER_ID =" & mlngDentEncId
    
    objDAS.ExecuteSQL strSQL
    
    strSQL = "Update DENTAL_PATIENT Set SBE_REQUIRE_FLAG=" & mintSBEReqFlag & " Where PATIENT_UNIT_ID=" & mlngPatUnitId
    
    objDAS.ExecuteSQL strSQL
    
    mblnDirty = False
    
FuncExit:
    Set objDAS = Nothing
    Set objRS = Nothing
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "frmSBE.UpdateDataBase", "", vbExclamation)
    GoTo FuncExit
End Sub

'*******************************************************************
'Function: RQ
'Author: Scott Foster
'Date: 11/28/01
'Returns: String
'Purpose: To replace single quotes with two single quotes for sql statements
'*******************************************************************
Private Function RQ(ByVal Value As String) As String
    
    'Replace regular single quotes
    RQ = Replace(Value, "'", "''")
    
    'Replace MS Word single quotes
    RQ = Replace(RQ, Chr(146), "''")
    
    'Replace MS Word open double quotes
    RQ = Replace(RQ, Chr(147), """")
   
    'Replace MS Word close double quotes
    RQ = Replace(RQ, Chr(148), """")
   
End Function

Private Sub IDENTAL_SBE_ClearSBE()
    mstrComments = Empty
    mintSBEReqFlag = 0
    mintSBE = 4
    
    mblnDirty = True
End Sub

Private Property Get IDENTAL_SBE_Comments() As String
    IDENTAL_SBE_Comments = mstrComments
End Property

Private Property Get IDENTAL_SBE_DentalEncounterId() As Long
    IDENTAL_SBE_DentalEncounterId = mlngDentEncId
End Property

Private Sub IDENTAL_SBE_GetUserInput(DentalEncounterId As String, PatientUnitId As String, Optional blnSBEForceDisp As Boolean = True)
     Dim stMousePointer As VBRUN.MousePointerConstants
     stMousePointer = Screen.MousePointer
     Screen.MousePointer = vbDefault
    
     mlngPatUnitId = PatientUnitId
     mlngDentEncId = DentalEncounterId
     
     frmSBE.SBE_Force_Display = blnSBEForceDisp
     frmSBE.Dental_Encounter_ID = DentalEncounterId
     frmSBE.Patient_Unit_ID = PatientUnitId
     
     If Trim(mstrComments) <> "" Then
        frmSBE.SBEComment.Text = mstrComments
     End If
     
     'frmSBE.DentalParent_Set Me
     cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
     frmSBE.InitSBEForm
     If frmSBE.SBERequireFlag_Get Then
        Load frmSBE
        frmSBE.Show vbModal
     End If
     With frmSBE
        mintSBE = .SBE
        mstrComments = .SBE_Comments
     End With
     
     If mintSBE = 0 Then
        mintSBEReqFlag = 0
     ElseIf mintSBE = 4 Then
        mintSBEReqFlag = 0
     Else
        mintSBEReqFlag = -1
     End If
     
     Unload frmSBE
     
     SaveSBE
     cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
     frmSBE.DentalParent_Clear
     frmSBE.SBEForceDisplay_Clear
     
     mblnDisplayRequiredSBE = False
     If blnSBEForceDisp Then mblnDisplayRequiredSBE = True
            
     Screen.MousePointer = stMousePointer


End Sub

Private Property Get IDENTAL_SBE_PatientUnitId() As Long
    IDENTAL_SBE_PatientUnitId = mlngPatUnitId
End Property

Private Property Get IDENTAL_SBE_RTFComments() As String
    Dim strReturn As String
    Dim blnDisplayFutureSBE As Boolean
    
    strReturn = "\viewkind4\uc1\pard\ul\b\f0\fs20 SBE Information:\ulnone\b0\par SBE: "

    strReturn = strReturn & GetSBEDesc(mintSBE)
    
    If mstrComments <> "" Then
        strReturn = strReturn & "\par SBE Comments: " & mstrComments & "\par "
    Else
        strReturn = strReturn & "\par "    'No SBE Comment
    End If
    If DisplayFutureSBE(mintSBE) Then
        strReturn = strReturn & "SBE Requirement: SBE Prophylaxis required for future " & _
            "encounters. \par"
    End If
    
    If DisplaySBE(mintSBE) Then
        IDENTAL_SBE_RTFComments = strReturn
    End If
End Property

Private Sub IDENTAL_SBE_SaveView()
    If mblnDirty Then
        SaveSBE
    End If
End Sub

Private Property Get IDENTAL_SBE_SBE() As Integer
    IDENTAL_SBE_SBE = mintSBE
End Property

Private Function GetSBEDesc(ByVal intSBE As Integer) As String
    Dim strReturn As String
    
    Select Case intSBE
        Case 1
            strReturn = "SBE prophylaxis taken as prescribed."
        Case 2
            strReturn = "SBE prophylaxis not taken as prescribed."
        Case 3
            strReturn = "SBE prophylaxis not required for this encounter."
        Case 4
            strReturn = "SBE prophylaxis not needed."
        Case Else
            strReturn = Empty
    End Select
    
    GetSBEDesc = strReturn
End Function

Private Function DisplayFutureSBE(ByVal intSBE As Integer) As Boolean
    Dim blnReturn As Boolean
    
    blnReturn = False
    If intSBE = 1 Or intSBE = 2 Or intSBE = 3 Then
        blnReturn = True
    End If
    
    DisplayFutureSBE = blnReturn
End Function
Private Function DisplaySBE(ByVal intSBE As Integer) As Boolean
    Dim blnReturn As Boolean
    If intSBE = 0 Then
        blnReturn = False
    Else
        blnReturn = True
    End If
    
    DisplaySBE = blnReturn
End Function
Private Property Get IDENTAL_SBE_SBERequiredFlag() As Boolean
    IDENTAL_SBE_SBERequiredFlag = mintSBEReqFlag
End Property
