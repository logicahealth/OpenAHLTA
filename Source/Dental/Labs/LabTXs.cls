VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabTXs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
'  Class Name: clsLabTXs
'
'================================================================================================
Option Explicit

'////////////////////  C L A S S   P R O P E R T I E S    S E C T I O N  ////////////////////

Private mcolLabTXs As Collection

'Methods of the LabTXs class.

'This method is called by GetLabCaseChildMembers method in clsLabCase
Public Function Add(ByRef oRst As Recordset, _
                     ByVal iChartType As Integer) As clsLabTX
                     
Dim oLabTX As New clsLabTX
Dim sTxADACodeID As String
Dim sToothLocation As String
Dim sSurfaceSetID As String

Dim sTXCodeID As String

Dim sKey As String
    
'error handling
On Error GoTo ErrHandler
oLabTX.FillFromRecordset oRst, iChartType

sKey = ""

If giLabCaseStatus = gcCreateNewLabCase Then
   'assign default SYSDATE value to TxDateOrdered
   'gsDefaultTXDateOrdered = basLibDAS.LibDAS_vSysDateTime("", gsDATE_FORMAT_VB)
   oLabTX.TXDateOrdered = gsDefaultTXDateOrdered
   oLabTX.RecordStatus = gcINSERT
Else
   'this is not CreateNewLabCase case, but new txs might be added
   If Len(CStr(oLabTX.LabTxID)) = 0 Then
      'new tx added, i.e. TXCodeID property is not populated with data yet
      'gsDefaultTXDateOrdered = basLibDAS.LibDAS_vSysDateTime("", gsDATE_FORMAT_VB)
      oLabTX.TXDateOrdered = gsDefaultTXDateOrdered
      oLabTX.RecordStatus = gcINSERT
   Else
      oLabTX.RecordStatus = gcStatusQuo
   End If
End If

'Add the LabTX object reference to the
'collection, using the TxADACodeID property as the key.
sTxADACodeID = oLabTX.TxADACodeID
sToothLocation = oLabTX.ToothLocation
sSurfaceSetID = oLabTX.SurfaceSetID

sTXCodeID = oLabTX.TxCodeID

'sKey = Trim$(sTxADACodeID) & Trim$(sToothLocation) & Trim$(sSurfaceSetID)
sKey = Trim$(sTXCodeID)

If Len(sKey) <> 0 Then
   mcolLabTXs.Add oLabTX, sKey
Else
   mcolLabTXs.Add oLabTX
End If

' Return a reference to the new LabTX object
Set Add = oLabTX

Exit Function

ErrHandler:

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTXs.Add", "Labs", vbCritical)

Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function Count() As Long
    Count = mcolLabTXs.Count
End Function

Public Sub Delete(ByVal vIndex As Variant)
    mcolLabTXs.Remove vIndex
End Sub

Public Function Item(ByVal vIndex As Variant) As clsLabTX
    Set Item = mcolLabTXs.Item(vIndex)
End Function

'NewEnum must return the IUnknown interface of a
'collection's enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolLabTXs.[_NewEnum]
End Function

Private Sub Class_Initialize()
On Error GoTo ErrHandler
'---set/initialize all data members

Set mcolLabTXs = New Collection

Exit Sub
ErrHandler:
   'MsgBox Err.Description, vbExclamation, "Error " & Err.Number & " in " & gsProcName
   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTXs.Initialize", "Labs", vbCritical)
End Sub

Public Function GetLabTXs(ByRef robjRS As Recordset, _
                           ByVal iChartType As Integer) As Boolean
''===========================================================================================
''  Description: This method gets the recordset of the lab treatments from database
''
''  Called By:   This function is called by GetLabCaseChildMembers method in
''               clsLabCase code
''
''  Argument(s):
''     robjRS - Contains recordset
''
''  Returns:
''         TRUE  - if recordset object successfully created and populated with data from db
''         FALSE - if failed
''
''  Notes:
''
''  Programmer:    Alek Samedov
''===========================================================================================

Dim sResult As String
Dim sSql As String

'--- error handling ---
On Error GoTo ErrHandler

GetLabTXs = True

Select Case iChartType
   Case gcGraphChart
      If giLabCaseStatus = gcCreateNewLabCase Then
                  
         sSql = "SELECT a.Tx_ADA_Code_ID, a.ADA_Category_Code, a.ADA_Definition, a.LabFlag, a.ADA_Code, " _
                  & "T1.Tooth_Name, T3.Surface_Set_ID, b.TX_CODE_ID, b.Status_ID " _
                  & "FROM TX_ADA_CODE_DEFINITION a, DENTAL_ENCOUNTER_TX b, DENTAL_ENCOUNTER DE, " _
                  & "DENTAL_ENCOUNTER_TXS_SS d, DENTAL_ENCOUNTER_TEETH T1, DENTAL_ENC_SURFACE_SET T3 " _
                  & "WHERE a.Tx_ADA_Code_ID = b.Tx_ADA_Code_ID AND " _
                  & "b.DENTAL_ENCOUNTER_ID = DE.DENTAL_ENCOUNTER_ID AND " _
                  & "DE.PATIENT_UNIT_ID = " & guCurEncInfo.PatientUnitID & " " _
                  & "AND b.TX_CODE_ID = d.TX_CODE_ID AND a.LabFlag = '" & glLabFlagOn & "' " _
                  & "AND T3.Dental_Encounter_Teeth_ID = T1.Dental_Encounter_Teeth_ID AND d.Surface_Set_ID = T3.Surface_Set_ID " _
                  & "AND (b.Status_ID = " & TxStatusEnum.txInProgress & " OR " & "b.Status_ID = " & TxStatusEnum.txCompleted & ") ORDER BY T1.Tooth_Name, a.ADA_Code "
      
      Else
         'lab case does exist in dental db and
         'lab tx records are getting from Lab_Tx_Specifications dynamic intersect table
                     
         sSql = "SELECT a.Tx_ADA_Code_ID, a.ADA_Category_Code, a.ADA_Definition, a.LabFlag, a.ADA_Code, " _
                  & "T1.Tooth_Name, T3.Surface_Set_ID, b.TX_CODE_ID, b.Status_ID " _
                  & "FROM TX_ADA_CODE_DEFINITION a, DENTAL_ENCOUNTER_TX b, DENTAL_ENCOUNTER DE, " _
                  & "DENTAL_ENCOUNTER_TXS_SS d, DENTAL_ENCOUNTER_TEETH T1, DENTAL_ENC_SURFACE_SET T3 " _
                  & "WHERE a.Tx_ADA_Code_ID = b.Tx_ADA_Code_ID " _
                  & "AND DE.DENTAL_ENCOUNTER_ID = b.DENTAL_ENCOUNTER_ID AND " _
                  & "DE.PATIENT_UNIT_ID = " & guCurEncInfo.PatientUnitID & " " _
                  & "AND b.TX_CODE_ID = d.TX_CODE_ID AND a.LabFlag = '" & glLabFlagOn & "' " _
                  & "AND T3.Dental_Encounter_Teeth_ID = T1.Dental_Encounter_Teeth_ID AND d.Surface_Set_ID = T3.Surface_Set_ID " _
                  & "AND (b.Status_ID = " & TxStatusEnum.txInProgress & " OR " & "b.Status_ID = " & TxStatusEnum.txCompleted & ") ORDER BY T1.Tooth_Name, a.ADA_Code "
      
      End If
   Case gcNonGraphChart
      If giLabCaseStatus = gcCreateNewLabCase Then
                  
         sSql = "SELECT a.Tx_ADA_Code_ID, a.ADA_Category_Code, a.ADA_Definition, a.LabFlag, a.ADA_Code, " _
                  & "T1.Location_Name, b.TX_CODE_ID, b.Status_ID " _
                  & "FROM TX_ADA_CODE_DEFINITION a, DENTAL_ENCOUNTER_TX b, " _
                  & "NOTE_TXS_LOCATION d, NOTE_LOCATION T1, DENTAL_ENCOUNTER DE " _
                  & "WHERE a.Tx_ADA_Code_ID = b.Tx_ADA_Code_ID " _
                  & "AND b.TX_CODE_ID = d.TX_CODE_ID AND b.Dental_Encounter_ID = DE.DENTAL_ENCOUNTER_ID " _
                  & "AND DE.PATIENT_UNIT_ID = " & guCurEncInfo.PatientUnitID & " AND a.LabFlag = '" & glLabFlagOn & "' " _
                  & "AND T1.Location_ID = d.Location_ID " _
                  & "AND (b.Status_ID = " & TxStatusEnum.txInProgress & " OR " & "b.Status_ID = " & TxStatusEnum.txCompleted & ") ORDER BY T1.Location_Name, a.ADA_Code "
      
      Else
         'lab case does exist in dental db and
         'lab tx records are getting from Lab_Tx_Specifications dynamic intersect table
                     
         sSql = "SELECT a.Tx_ADA_Code_ID, a.ADA_Category_Code, a.ADA_Definition, a.LabFlag, a.ADA_Code, " _
                  & "T1.Location_Name, b.TX_CODE_ID, b.Status_ID " _
                  & "FROM TX_ADA_CODE_DEFINITION a, DENTAL_ENCOUNTER_TX b, " _
                  & "NOTE_TXS_LOCATION d, NOTE_LOCATION T1, DENTAL_ENCOUNTER DE " _
                  & "WHERE a.Tx_ADA_Code_ID = b.Tx_ADA_Code_ID " _
                  & "AND b.TX_CODE_ID = d.TX_CODE_ID AND b.Dental_Encounter_ID = DE.DENTAL_ENCOUNTER_ID " _
                  & "AND DE.PATIENT_UNIT_ID = " & guCurEncInfo.PatientUnitID & " AND a.LabFlag = '" & glLabFlagOn & "' " _
                  & "AND T1.Location_ID = d.Location_ID " _
                  & "AND (b.Status_ID = " & TxStatusEnum.txInProgress & " OR " & "b.Status_ID = " & TxStatusEnum.txCompleted & ") ORDER BY T1.Location_Name, a.ADA_Code "
      
      End If
   
End Select

sResult = LabDataProcessing(sSql, robjRS)
If sResult = gcRowsSelected Then
   GetLabTXs = True
Else
   GetLabTXs = False
End If

Exit Function

ErrHandler:
GetLabTXs = False

   Call gobjShared.ShowVBError(Err.Number, Err.Description, _
                "clsLabTXs.GetLabTXs", "Labs", vbCritical)

End Function

Private Sub Class_Terminate()
   Set mcolLabTXs = Nothing
End Sub
