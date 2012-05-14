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
Attribute VB_Name = "ProxyCDRPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Handles all direct database access within A/P.  The database can vary depend on
Option Explicit

Private Const LAB_SEQ As String = "oe_lab_data_id_sequence"
Private Const PI_SEQ As String = "pi_data_data_id_sequence"
Private Const MED_SEQ As String = "oe_rx_data_id_sequence"
Private Const RAD_SEQ As String = "oe_rad_data_id_sequence"

Private Const SQLServerDate As String = "GetDate()"
Private Const OracleDate As String = "sysdate"
Private Const MySQLDate As String = "sysdate()"

Private mDAS As ICHCSII_DAS

Public Function AddNewLab(ByRef Lab As OrderLab) As String

    'In Theater and ITT GEM_OE does the save before the order is added to A/P.
    'Orders check for a value returned for success so set to anything.
    AddNewLab = "Fake Save"
  
End Function

Public Function AddNewMed(ByRef Med As OrderMed) As String

    'In Theater and ITT GEM_OE does the save before the order is added to A/P.
    'Orders check for a value returned for success so set to anything.
    AddNewMed = "Fake Save"
    Exit Function
  
End Function

Public Function AddNewPI(ByVal SnoID As String) As String
Dim SQLStmt As String
Dim DataID As String
Dim DBDate As String
Dim FacilityID As String
Dim EncounterID As String
Dim Success As Boolean
        
  On Error GoTo ErrHandler
        
  FacilityID = gBB.CurrentEncounter.FacilityNCID
  EncounterID = gBB.CurrentEncounter.EncounterID
        
  DataID = GetNewID(PI_SEQ)
  DBDate = GetDatabaseDateValue
        
  SQLStmt = "INSERT INTO pi_data (pi_data_id, enc_fac_ncid, enc_num, SNOID, pi_time) VALUES  (" & _
            DataID & "," & FacilityID & ",'" & EncounterID & "'" & "," & _
            SnoID & ", " & DBDate & ")"
                        
  Success = ExecuteSQL(SQLStmt)
                        
  If Success Then AddNewPI = DataID
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.AddNewPI", "A/P", vbExclamation
  
End Function
'scr 56524 KMonroe 11/03/04
Public Function AddNewInPat() As String
        
    'In Theater and ITT GEM_OE does the save before the order is added to A/P.
    'Orders check for a value returned for success so set to anything.
    AddNewInPat = "Fake Save"
    Exit Function
  
End Function

Public Function AddNewRad(ByRef Rad As OrderRad) As String

    'In Theater and ITT GEM_OE does the save before the order is added to A/P.
    'Orders check for a value returned for success so set to anything.
    AddNewRad = "Fake Save"
    Exit Function

End Function

Private Function CommonWhereClause() As String
Dim FacilityID As String
Dim EncounterID As String

  On Error GoTo ErrHandler
  
  FacilityID = gBB.CurrentEncounter.FacilityNCID
  EncounterID = gBB.CurrentEncounter.EncounterID

  CommonWhereClause = "WHERE enc_fac_ncid = " & FacilityID & _
                      " and enc_num = '" & EncounterID & "' "
                    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.CommonWhereClause", "A/P", vbExclamation
                    
End Function
Private Function OtherWhereClause() As String
Dim FacilityID As String
Dim EncounterID As String

  On Error GoTo ErrHandler
  
  FacilityID = gBB.CurrentEncounter.FacilityNCID
  EncounterID = gBB.CurrentEncounter.EncounterID

  '--- SCR 64289   Sherry Wang   1/13/05
  OtherWhereClause = "WHERE facility_ncid = " & FacilityID & _
                      " and encounter_number = '" & EncounterID & "' " & _
                      " and ORDER_STATUS <> 6"
                    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.OtherWhereClause", "A/P", vbExclamation
                    
End Function

'scr 21524  7/17/02  Phil Crowder  see LoadRads and LoadLabs methods
Private Function ConvertStatusToText(ByVal Status As Long) As String
  On Error Resume Next
  
  Select Case Status
    Case 0
      ConvertStatusToText = "Preliminary"
    Case 1
      ConvertStatusToText = "Pending"
    Case 2
      ConvertStatusToText = "Completed Unaddressed"
    Case 3
      ConvertStatusToText = "Completed Addressed"
  End Select

End Function

Public Function DeleteLab(ByVal DataID As String) As Boolean
      
    'In Theater and ITT, GEM_OE does the delete from the database.
    DeleteLab = True
    Exit Function

End Function
'scr 56524 KMonroe 11/03/04
Public Function DeleteInPat(ByVal DataID As String) As Boolean
     
    'In Theater and ITT, GEM_OE does the delete from the database.
    DeleteInPat = True
    Exit Function

End Function

Public Function DeleteMed(ByVal DataID As String) As Boolean
      
    'In Theater and ITT, GEM_OE does the delete from the database.
    DeleteMed = True

End Function

Public Function DeletePI(ByVal DataID As String) As Boolean
Dim SQLStmt As String
      
  On Error GoTo ErrHandler
      
  SQLStmt = "UPDATE pi_data SET enc_fac_ncid = null, enc_num = null " & _
            CommonWhereClause & "and pi_data_id = '" & DataID & "'"
                    

  DeletePI = ExecuteSQL(SQLStmt)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.DeletePI", "A/P", vbExclamation

End Function

Public Function DeleteRad(ByVal DataID As String) As Boolean
    
    'In Theater and ITT, GEM_OE does the delete from the database.
    DeleteRad = True

End Function

Private Function ExecuteInsertSP(ByRef oCmd As CHCSII_Command.CHCSII_Cmd, ByRef DataID As String) As Boolean
  
On Error GoTo ErrHandler
  
Dim oRS As ADODB.Recordset

   InitConn
   
   Set oRS = mDAS.OpenRecordsetSP(oCmd)
   
   If Not oRS Is Nothing Then
      If Not oRS.BOF And Not oRS.EOF Then
         If oRS(0).Value > 0 Then
            DataID = oRS(0)
            ExecuteInsertSP = True
         Else
            Err.Raise 911, , "No Data Id Returned.  Outcome was: " & oRS(1).Value
         End If
      End If
   End If
      
  
ErrHandler:
   'Clean Up Objects
   Set oRS = Nothing

   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description & vbCrLf & " SP: " & oCmd.SPName, "ProxyCDRPlus.ExecuteInsertSP", "A/P", vbExclamation
  
End Function


Private Function ExecuteSQL(ByVal SQLStmt As String) As Boolean
  
  On Error GoTo ErrHandler
  
  InitConn
  mDAS.ExecuteSQL SQLStmt
  
  ExecuteSQL = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description & " SQL: " & SQLStmt, "ProxyCDRPlus.ExecuteSQL", "A/P", vbExclamation
  Exit Function
  Resume
End Function

Private Function FilterNumericNulls(ByVal Value As Variant) As Variant
  If Not IsNumeric(Value) Then
    FilterNumericNulls = 0
  Else
    FilterNumericNulls = Value
  End If
End Function

Private Function GetDatabaseDateValue() As String
  If gShared.UseMySQLDB Then
    GetDatabaseDateValue = MySQLDate
  ElseIf gShared.UseSQLServer Then
    GetDatabaseDateValue = SQLServerDate
  Else
    GetDatabaseDateValue = OracleDate
  End If
End Function

Private Function GetNewID(ByVal Sequence As String) As String
Dim rs As Recordset

  On Error GoTo ErrHandler

  InitConn
  
  'JT SCR 48267 - use the DAS object for ITT Mode
  If gShared.IsAppMode(modeCDR) And Not gShared.IsAppMode(modeITT) Then
    Set rs = mDAS.OpenRecordset("SELECT " & Sequence & ".nextval FROM dual")
    GetNewID = rs!NEXTVAL
    rs.Close
  Else
    GetNewID = mDAS.GetNextID()
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.GetNewID", "A/P", vbExclamation
  
End Function

Private Sub InitConn()
Dim Connection As CHCSII_CONN.Conn

  On Error GoTo ErrHandler
  
  If Not (mDAS Is Nothing) Then Exit Sub
  
  Set Connection = New CHCSII_CONN.Conn
  Set mDAS = Connection.CHCSII_DAS
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.InitConn", "A/P", vbExclamation

End Sub

'Codes, Short Decription, Help
Public Function ListOfAllModifiers() As Recordset
Dim SQLStmt As String
Dim rs As Recordset

  InitConn
   
  SQLStmt = "SELECT Modifier_Code, Short_Description, Long_Description " & _
            "FROM VPROCEDURE_MODIFIERS WHERE FACILITY_NCID = " & gBB.Logon.FacilityNCID
              
  Set rs = mDAS.OpenRecordset(SQLStmt)

  Set ListOfAllModifiers = rs
  
End Function

'Array of Codes, Short Decription, Help
Public Function ListOfModifiers(ByVal Code As String) As Variant
Dim SQLStmt As String
Dim rs As Recordset

  InitConn
   
  SQLStmt = "SELECT Modifier_Code, Short_Description, Long_Description " & _
            "FROM VPROCEDURE_MODIFIERS WHERE Facility_NCID = " & gBB.Logon.FacilityNCID & " and " & _
            "CPTCode_Range_Begin <='" & Code & "' AND CPTCode_Range_End >= '" & Code & "'"
              
  Set rs = mDAS.OpenRecordset(SQLStmt)

  If Not rs.BOF And Not rs.EOF Then
    ListOfModifiers = rs.GetRows
  End If
  
End Function

Private Function LoadLab(ByRef rsFields As Fields) As iOrder
Dim AnPOrder As iOrder
Dim Lab As OrderLab
    
  On Error GoTo ErrHandler
    
  Set AnPOrder = New OrderLab
    
  AnPOrder.Init rsFields!oe_lab_data_id, rsFields!order_name
    
  Set Lab = AnPOrder
  
  With Lab
    .OrderID = rsFields!order_id
    .OrderIEN = rsFields!order_ien
    .Comment = rsFields!lab_comment & vbNullString
    .CollectionMethod = rsFields!collect_method & vbNullString
    .ProcessPriority = rsFields!process_priority
    .LabCollectionPriority = rsFields!lab_collect_priority
    .SpecimenName = rsFields!specimen_name & vbNullString
    .SpecimenIEN = rsFields!specimen_ien & vbNullString
    .SpecimenTogography = rsFields!specimen_topography & vbNullString
    .SpecimenTogographyIEN = rsFields!specimen_topography_ien & vbNullString
    .SpecimenAnswerCode = rsFields!specimen_answer_code & vbNullString
    .SpecimenAnswerText = rsFields!specimen_answer_text & vbNullString
    .SpecimenContainer = rsFields!specimen_capcolor_or_container & vbNullString
    .TimingOneOrContinued = FilterNumericNulls(rsFields!timing_one_x_or_continued)
    .TimingAdminTime = rsFields!timing_admin_time & vbNullString
    .TimingDuration = FilterNumericNulls(rsFields!timing_duration)
    .TimingFrequence = rsFields!timing_frequency & vbNullString
    .TimingStartDate = rsFields!timing_start_date
    .UnexpandedTime = rsFields!unexpanded_time & vbNullString
    .OrderedBy = rsFields!ORDERED_BY & vbNullString
    .Provider = rsFields!Provider & vbNullString
    .TimeOrdered = rsFields!time_ordered
    
    'scr 21524  7/17/02  Phil Crowder
    If gShared.IsAppMode(modeTheater) Then
      .Results = rsFields!Results & vbNullString
      .Status = ConvertStatusToText(rsFields!Status)
    End If
  End With

  Set LoadLab = AnPOrder
  
  Exit Function
  
ErrHandler:
  Set LoadLab = AnPOrder
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadLab", "A/P", vbExclamation
  
End Function

Public Function LoadLabs() As Collection
Dim SQLStmt As String
Dim rs As Recordset
Dim AnPOrders As Collection
Dim AnPOrder As iOrder

  On Error GoTo ErrHandler
   
  Set AnPOrders = New Collection
  
  InitConn
   
  If gShared.IsAppMode(modeTheater) Then 'Jane Shen 12/08/2004, only load root orders
    SQLStmt = "SELECT * FROM oe_lab " & CommonWhereClause & " And IS_ORDER_OCCURRENCE = 0 and status <> 6 "
  Else
    SQLStmt = "SELECT * FROM oe_lab " & CommonWhereClause
  End If
  
  Set rs = mDAS.OpenRecordset(SQLStmt)

  Do While Not rs.EOF
    Set AnPOrder = LoadLab(rs.Fields)
    AnPOrders.Add AnPOrder, AnPOrder.Key
    
    Set AnPOrder = Nothing
    
    rs.MoveNext
  Loop

  rs.Close
  
  Set LoadLabs = AnPOrders
  
  Exit Function
  
ErrHandler:
  Set LoadLabs = AnPOrders
  If Not (rs Is Nothing) Then rs.Close
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadLabs", "A/P", vbExclamation

End Function

Private Function LoadMed(ByRef rsFields As Fields) As iOrder
Dim AnPOrder As iOrder
Dim Med As OrderMed
Dim rs As ADODB.Recordset
Dim SQLStmt As String

Dim OverrideReasonID As String

  On Error GoTo ErrHandler
  
  Set AnPOrder = New OrderMed
    
  AnPOrder.Init rsFields!oe_rx_data_id, rsFields!order_name
    
  Set Med = AnPOrder
    
  With Med
    'SCR 52839 - Updated to handle null values in some numeric fields
    .StartDate = rsFields!ORDER_DATE
    .Comment = rsFields!rx_comment & vbNullString
    .SigCode = rsFields!sig_code & vbNullString
    .Quantity = FilterNumericNulls(rsFields!Quantity)
    .Refills = FilterNumericNulls(rsFields!Refills)
    .PharmacySent = rsFields!pharmacy_name_sent & vbNullString
    .PharmacySentIEN = rsFields!pharmacy_ien_sent & vbNullString
    .ChildCap = FilterNumericNulls(rsFields!child_cap)
    .OrderID = rsFields!order_id
    .OrderIEN = FilterNumericNulls(rsFields!order_ien)
    .OrderedBy = rsFields!ORDERED_BY & vbNullString
    .Provider = rsFields!Provider & vbNullString
    .TimeOrdered = rsFields!time_ordered
  End With

  'SCR 59827 - Inpt II Rx Checking-13 Warning Display on SF600
  '12/22/04 Frank
  If LenB(gShared.CmdLineSwitches("INPATIENT2")) > 0 Then
    
    Call InitConn
  
    SQLStmt = "Select Override_comment, Override_Reason from Medications where FDB_MEDID  = " & Med.OrderIEN & " and ordernumber = '" & Med.OrderID & "'"
    Set rs = mDAS.OpenRecordset(SQLStmt)
    If Not (rs.BOF And rs.EOF) Then
      Med.OverrideComment = rs.Fields("Override_comment").Value & vbNullString
      OverrideReasonID = "0" & rs.Fields("Override_Reason").Value
    End If
    
    If CLng("0" & OverrideReasonID) > 0 Then
        SQLStmt = "Select Reason_Description from Override_Reason where ID = " & OverrideReasonID
        Set rs = mDAS.OpenRecordset(SQLStmt)
        If Not (rs.BOF And rs.EOF) Then
            Med.OverrideReason = rs.Fields("Reason_Description").Value & vbNullString
        End If
    End If
    
  End If

  Set LoadMed = AnPOrder
  
  Exit Function
  
ErrHandler:
  Set LoadMed = AnPOrder
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadMed", "A/P", vbExclamation
  
End Function

Public Function LoadMeds() As Collection

    Dim SQLStmt As String
    Dim rs As Recordset
    Dim AnPOrders As Collection
    Dim AnPOrder As iOrder
 
    Const NCID_ORDERSTATUS_DISCONTINUED = "1524"
    Const NCID_ORDERSTATUS_ACTIVE = "1024"
    Const NCID_ORDERSTATUS_COMPLETE = "9"
    Const NCID_PARTIAL = "14516644"
    Const NCID_ORDERSTATUS_SCHEDULED = "1555"
    Const NCID_ONHOLD = "163002"
    Const NCID_ORDERSTATUS_EXPIRED = "206497"
    Const NCID_EXPIRED = "14516641"


    On Error GoTo ErrHandler
   
    Set AnPOrders = New Collection
    
    InitConn
    
    If gShared.IsAppMode(modeTheater) Then
        'SCR 76469 Don't load child med orders.
        SQLStmt = "SELECT * FROM oe_rx " & CommonWhereClause & " and CHILD_MED_IEN is NULL"
    Else
        SQLStmt = "SELECT * FROM oe_rx " & CommonWhereClause
    End If
              
    'SCR 49807 - In theater we need an additional filter to not include discontinued meds
    If gShared.IsAppMode(modeTheater) Then
        SQLStmt = SQLStmt & " AND status_ncid <> " & NCID_ORDERSTATUS_DISCONTINUED
    End If

    Set rs = mDAS.OpenRecordset(SQLStmt)
    
    Do While Not rs.EOF
      If getMultiOrderCount(rs.Fields("ORDER_ID").Value) > 1 Then
          If rs.Fields("STATUS_NCID") = NCID_PARTIAL Then
              '*** SKIP PENDING REFILL MED ORDER
          Else
              Set AnPOrder = LoadMed(rs.Fields)
              AnPOrders.Add AnPOrder, AnPOrder.Key
          End If
      Else
          Set AnPOrder = LoadMed(rs.Fields)
          AnPOrders.Add AnPOrder, AnPOrder.Key
      End If
      Set AnPOrder = Nothing
      
      rs.MoveNext
    Loop
    
    rs.Close
  
    Set LoadMeds = AnPOrders
  
  Exit Function
  
ErrHandler:
  Set LoadMeds = AnPOrders
  If Not (rs Is Nothing) Then rs.Close
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadMeds", "A/P", vbExclamation

End Function


Public Function getMultiOrderCount(ByVal order_id As String) As Integer
    
    Dim strSQL As String
    Dim rs     As Recordset

    On Error GoTo ErrHandler

    strSQL = "Select count(*) from oe_rx where order_id = '" & order_id & "'"
    Set rs = mDAS.OpenRecordset(strSQL)
    getMultiOrderCount = rs.Fields.Item(0).Value

    rs.Close
    Exit Function

ErrHandler:
    gShared.ShowVBError Err.Number, Err.Description, "Global.getMultiOrderCount", "sharedGlobal", vbExclamation
    
End Function


Private Function LoadPI(ByRef rsFields As Fields, ByVal Convert As IConvert) As iOrder
Dim AnPOrder As iOrder
Dim PI As OrderPatientInstruc
Dim Name As String

  On Error GoTo ErrHandler
  
  Set AnPOrder = New OrderPatientInstruc
  
  Name = Convert.SnoDescription(rsFields("SnoID"))

  AnPOrder.Init rsFields!pi_data_id, Name
  
  Set PI = AnPOrder
  PI.SnoID = rsFields("SnoID")
  
  Set LoadPI = AnPOrder
  
  Exit Function
  
ErrHandler:
  Set LoadPI = AnPOrder
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadPI", "A/P", vbExclamation
  Exit Function
  Resume
End Function

Public Function LoadPIs() As Collection
Dim SQLStmt As String
Dim rs As Recordset
Dim AnPOrders As Collection
Dim AnPOrder As iOrder
Dim Convert As IConvert

  On Error GoTo ErrHandler
     
  Set AnPOrders = New Collection
  
  InitConn

  SQLStmt = "SELECT pi_data_id, snoid FROM pi_data " & CommonWhereClause
               
  Set rs = mDAS.OpenRecordset(SQLStmt)

  If Not rs.EOF Then Set Convert = New IConvert
  
  Do While Not rs.EOF
    Set AnPOrder = LoadPI(rs.Fields, Convert)
    AnPOrders.Add AnPOrder, AnPOrder.Key
    
    Set AnPOrder = Nothing
    
    rs.MoveNext
  Loop

  rs.Close
  
  Set LoadPIs = AnPOrders

  Exit Function
  
ErrHandler:
  Set LoadPIs = AnPOrders
  If Not (rs Is Nothing) Then rs.Close
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadPIs", "A/P", vbExclamation

End Function
'scr 56524 KMonroe 11/03/04
Public Function LoadInPats() As Collection
Dim SQLStmt As String
Dim rs As Recordset
Dim AnPOrders As Collection
Dim AnPOrder As iOrder
Dim Convert As IConvert

  On Error GoTo ErrHandler
     
  Set AnPOrders = New Collection
  
  InitConn

  SQLStmt = "SELECT * FROM oe_other " & OtherWhereClause
               
  Set rs = mDAS.OpenRecordset(SQLStmt)
 
  Do While Not rs.EOF
    Set AnPOrder = LoadInPat(rs.Fields)
    AnPOrders.Add AnPOrder, AnPOrder.Key
    
    Set AnPOrder = Nothing
    
    rs.MoveNext
  Loop

  rs.Close
  
  Set LoadInPats = AnPOrders

  Exit Function
  
ErrHandler:
  Set LoadInPats = AnPOrders
  If Not (rs Is Nothing) Then rs.Close
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadInPats", "A/P", vbExclamation

End Function
'scr 56524 KMonroe 11/03/04
Private Function LoadInPat(ByRef rsFields As Fields) As iOrder
Dim AnPOrder As iOrder
Dim InPat As OrderInPatient
Dim Name As String

  On Error GoTo ErrHandler
  
  Set AnPOrder = New OrderInPatient
  
  AnPOrder.Init rsFields!OE_OTR_DATA_ID, rsFields!ORDER_TYPE & ", " & rsFields!ORDER_TYPE_DESC
    
  Set InPat = AnPOrder
     
  With InPat
    .OrderDate = rsFields!ORDER_DATE
    .Comment = rsFields!ORDER_COMMENTS & vbNullString
    .OrderedBy = rsFields!ORDERED_BY & vbNullString
    .OrderID = rsFields!OE_OTR_DATA_ID
    .FacilityNCID = rsFields!FACILITY_NCID & vbNullString
    .OrderName = rsFields!ORDER_TYPE & ", " & rsFields!ORDER_TYPE_DESC & vbNullString
    .OrderType = rsFields!ORDER_TYPE & vbNullString
    .OrderStatus = rsFields!ORDER_STATUS
    If Not IsNull(rsFields!NURSE_ACK_FLAG) Then
        .NurseACK = CBool(rsFields!NURSE_ACK_FLAG)
    End If
    .NurseAckName = rsFields!NURSE_ACK_NAME & vbNullString
    If Not IsNull(rsFields!NURSE_ACK_DATE) Then
        .NurseAckDate = rsFields!NURSE_ACK_DATE
    End If
    
    If Not IsNull(rsFields!SIGN_FLAG) Then
        .Signed = CBool(rsFields!SIGN_FLAG)
    End If
  End With

  Set LoadInPat = AnPOrder
  
  Exit Function
  
ErrHandler:
  Set LoadInPat = AnPOrder
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadInPat", "A/P", vbExclamation
  
End Function

Private Function LoadRad(ByRef rsFields As Fields) As iOrder
Dim AnPOrder As iOrder
Dim Rad As OrderRad
    
  On Error GoTo ErrHandler
  
  Set AnPOrder = New OrderRad
    
  AnPOrder.Init rsFields!oe_rad_data_id, rsFields!order_name
    
  Set Rad = AnPOrder
  
  With Rad
    .OrderID = rsFields!order_id
    .OrderIEN = rsFields!order_ien
    .Comment = rsFields!rad_comment & vbNullString
    .ClinicalImpression = rsFields!clinical_impression_text & vbNullString
    .ProcessPriority = rsFields!process_priority
    .TransportionMode = rsFields!transport_mode & vbNullString
    .WetReadingWanted = FilterNumericNulls(rsFields!wet_reading_wanted)
    .TimingOneOrContinued = FilterNumericNulls(rsFields!timing_one_x_or_continued)
    .TimingAdminTime = rsFields!timing_admin_time & vbNullString
    .TimingDuration = FilterNumericNulls(rsFields!timing_duration)
    .TimingFrequence = rsFields!timing_frequency & vbNullString
    .TimingStartDate = rsFields!timing_start_date
    .OrderedBy = rsFields!ORDERED_BY & vbNullString
    .Provider = rsFields!Provider & vbNullString
    .TimeOrdered = rsFields!time_ordered
    
    'scr 21524  7/17/02  Phil Crowder
    If gShared.IsAppMode(modeTheater) Then
      .Results = rsFields!Results & vbNullString
      .Status = ConvertStatusToText(rsFields!Status)
    End If
  End With

  Set LoadRad = AnPOrder
  
  Exit Function
  
ErrHandler:
  Set LoadRad = AnPOrder
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadRad", "A/P", vbExclamation
  
End Function

Public Function LoadRads() As Collection
Dim SQLStmt As String
Dim rs As Recordset
Dim AnPOrders As Collection
Dim AnPOrder As iOrder
 
  On Error GoTo ErrHandler
   
  Set AnPOrders = New Collection
  
  InitConn
  
  If gShared.IsAppMode(modeTheater) Then 'Jane Shen 12/08/2004, only load root orders
    SQLStmt = "select * from oe_rad " & CommonWhereClause & " And IS_ORDER_OCCURRENCE = 0 and status <> 6"
  Else
    SQLStmt = "select * from oe_rad " & CommonWhereClause
  End If
  
  Set rs = mDAS.OpenRecordset(SQLStmt)

  Do While Not rs.EOF
    Set AnPOrder = LoadRad(rs.Fields)
    AnPOrders.Add AnPOrder, AnPOrder.Key
    
    Set AnPOrder = Nothing
    
    rs.MoveNext
  Loop

  Set LoadRads = AnPOrders

  Exit Function
  
ErrHandler:
  Set LoadRads = AnPOrders
  If Not (rs Is Nothing) Then rs.Close
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.LoadRads", "A/P", vbExclamation

End Function

'SCR 19064  11/28/01  Phil Crowder
'Added MaxLength parameter to do left trim function in this function so that we
'can account for the possiblity of the trimed string ending in a single quote.
Private Function SQLQuote(ByVal sTxt As String, Optional ByVal MaxLength As Long = 0) As String
    sTxt = Replace(Replace(sTxt, Chr$(34), Chr$(34) & Chr$(34)), "'", "''")
    
    If MaxLength > 0 Then
      sTxt = Left$(sTxt, MaxLength)
      
      'Make sure that the string does not end in a single quote
      If Right$(sTxt, 1) = "'" Then
        If Right$(sTxt, 2) <> "''" Then
          sTxt = Left$(sTxt, Len(sTxt) - 1)
        End If
      End If
    End If
    
    SQLQuote = "'" & sTxt & "'"
    
End Function

Public Function UpdateLab(ByRef Lab As OrderLab) As Boolean
    
    'In Theater and ITT GEM_OE does the save so there is nothing for AP to save
    UpdateLab = True
  
End Function

Public Function UpdateMed(ByRef Med As OrderMed) As Boolean
    
    'In Theater and ITT GEM_OE does the save so there is nothing for AP to save
    UpdateMed = True
  
End Function

Public Function UpdateRad(ByRef Rad As OrderRad) As Boolean
    
    'In Theater and ITT GEM_OE does the save so there is nothing for AP to save
    UpdateRad = True
  
End Function
Public Function UpdateInPat(ByRef InPatOrder As OrderInPatient) As Boolean

    'In Theater and ITT GEM_OE does the save so there is nothing for AP to save
    UpdateInPat = True

End Function


