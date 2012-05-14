VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITT_SQLOPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_SQL.ISqlOps
Implements ISec_SQL.ISecSqlOps
Implements ICHCSII_SQL.ISqlOpsEx
Implements ICHCSII_SQL.ISQLDirect

Private sConnString As String

Private mobjShared As CWShared

Private mvarRecordSet As ADODB.Recordset
Public UserNCID As String
Private mvarSQLStatement As String
Private mNoAudit As Collection

Private mFacilityNCID As String  '- for backfill of providers
Private mClinicNCID   As String

#If DEBUGON Then
  Public DL                             As cDebugLog
  Private Const Exe                      As String = "GEMS_SQLOPS"
  Private Const Module                  As String = "GEMS_SQLOPS CLS"
#End If

Private lTimeBias As Long
Private TZInfo As TIME_ZONE_INFORMATION


Private Function ISqlOps_ExecuteSP(Parms As CHCSII_Cmd) As Variant

End Function

Private Function ISqlOps_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.eSeq = 0&) As Variant
'-- pulls the next value from the MainSeq sequence
    On Error GoTo ErrHandler
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim sSeqName As String
    Dim lSeqNum As Long
    
    sSeqName = SeqEnumToString(eSeqName)
    
    If eSeqName = esMAINSEQ Then
        sSeqName = "PROVIDER_NCID"
    End If
    
    oRs.Open "Select Value from Sequences where seq_type = '" & sSeqName & "'", goITTOPSCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not oRs.EOF Then
        lSeqNum = Val(oRs(0))
    End If
    
    goITTOPSCnn.Execute "update Sequences set value = " & lSeqNum + 1 & " where seq_type = '" & sSeqName & "'"
    ISqlOps_GetNextID = lSeqNum

ErrHandler:
    If oRs.State > 0 Then oRs.Close
    Set oRs = Nothing

End Function

'<SCR 22426>
'Appliacle only to CHCSII_SQL
Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNcid As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean

End Function
Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNcid As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean

End Function
'</SCR 22426>

Private Function ISqlOps_ValidateFacility(FacilityNcid As String, ByVal CreateIfNotFound As Boolean, Optional ClinicNCID As String = "") As Boolean

End Function

Private Function ISqlOpsEx_AppendAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, append the byte array as chunks
    Dim lsize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim ChunkSize   As Long
    Dim index       As Integer
'    Dim oZLib       As ZLib
    Dim lResult     As Long
    Dim sTxt As String
    
    If IsNumeric(vColumn) Then
        index = vColumn
    Else
        index = ISqlOps_ColumnIndex(CStr(vColumn))
    End If
    
    
    lsize = LenB(vText)
'- should we compress it?
'    If lsize > 4000 Then
'    '-- yes
'        Set oZLib = New ZLib
'        sTxt = vText
'
'        lResult = oZLib.CompressString(sTxt)
'        Set oZLib = Nothing
'    '-- label it as such
'        vText = Format$(lsize, "ZLIB000000000") & sTxt
'    End If

    '-- pick a chunk size
    ChunkSize = 2000
    If LenB(vText) > 10000 Then
        ChunkSize = LenB(vText) \ 10
        '- Make it EVEN (2-bytes per char, remember?)
        If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
        If ChunkSize > 20000 Then ChunkSize = 20000
    End If

    
    lsize = LenB(vText)
    loffset = 0 ' Reset offset
    
    Do While loffset < lsize
       varChunk = LeftB(RightB(vText, lsize - loffset), ChunkSize)
        mvarRecordSet.Fields(index).AppendChunk varChunk
       loffset = loffset + ChunkSize
    Loop
    
    
    ISqlOpsEx_AppendAsChunk = True
    Exit Function
ErrHandler:

End Function

Private Sub ISqlOpsEx_BeginTrans():    goITTOPSCnn.BeginTrans:    End Sub
Private Sub ISqlOpsEx_CommitTrans():   goITTOPSCnn.CommitTrans:   End Sub


Private Function ISqlOpsEx_GetAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, get the data as chunks and return in the byte array
'- for the column indicated, append the byte array as chunks
    Dim lsize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim index       As Integer
    Dim ChunkSize   As Long
'    Dim oZLib       As ZLib
    Dim lResult     As Long
    Dim sTxt        As String
    
    If IsNumeric(vColumn) Then
        index = vColumn
    Else
        index = ISqlOps_ColumnIndex(CStr(vColumn))
    End If
    
       
   
    lsize = mvarRecordSet.Fields(index).ActualSize
    loffset = 0 ' Reset offset
    vText = ""
    
    '-- set chunksize
    ChunkSize = 2000
    If lsize > 10000 Then
        ChunkSize = lsize \ 10
        '- Make it EVEN (2-bytes per char, remember?)
        If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
        If ChunkSize > 20000 Then ChunkSize = 20000
    End If
    
    Do While loffset < lsize
       varChunk = mvarRecordSet.Fields(index).GetChunk(ChunkSize)
       vText = vText & varChunk
       loffset = loffset + ChunkSize
    Loop
        
    '-- see if we compressed it
'    If Left(vText, 4) = "ZLIB" Then
'        lsize = Val(Mid(vText, 5, 9))
'        vText = Mid(vText, 14)
'        sTxt = vText
'        Set oZLib = New ZLib
'        lResult = oZLib.DecompressString(sTxt, lsize)
'        Set oZLib = Nothing
'    End If
            

End Function

Private Sub ISqlOpsEx_RollBackTrans(): goITTOPSCnn.RollbackTrans: End Sub


Private Function ISecSqlOps_Connect() As Boolean

    On Error GoTo ErrHandler
    
    ISecSqlOps_Connect = True
    
Exit Function
    
'    On Error Resume Next
'    goITTOPSCnn.Close
'    'Set goITTOPSCnn = DESec.Cnn
'
'    sConnString = GEMS_Sec_Conn
'    goITTOPSCnn.ConnectionString = GetConnectionString(edb_Oracle)
'    goITTOPSCnn.ConnectionString = ReplaceParameter(goITTOPSCnn.ConnectionString, "USER ID=", "secadmin")
'    goITTOPSCnn.ConnectionString = ReplaceParameter(goITTOPSCnn.ConnectionString, "PASSWORD=", "secadmin")
'
'    goITTOPSCnn.Open
'    If goITTOPSCnn.State = 1 Then
'        Call SetDBDateFormat(goITTOPSCnn)
'    End If
'
'    Set mvarRecordSet = New ADODB.Recordset
'    ISecSqlOps_Connect = True
    
ErrHandler:
End Function

Public Function SecDisConnect() As Boolean

    On Error GoTo ErrHandler
    
    SecDisConnect = True

Exit Function

'    On Error Resume Next
'    goITTOPSCnn.Close
'
'    sConnString = GEMS_Conn
'    goITTOPSCnn.ConnectionString = GetConnectionString(edb_Oracle)
'    goITTOPSCnn.Open
'    If goITTOPSCnn.State = 1 Then
'        Call SetDBDateFormat(goITTOPSCnn)
'    End If
'
'    Set mvarRecordSet = New ADODB.Recordset
'
'    SecDisConnect = True
    
ErrHandler:
End Function
Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant
    Dim lRecs As Long     '- number of records affected
    
    oAdoCmd.ActiveConnection = goITTOPSCnn
    
    Select Case oAdoCmd.CommandType
    Case adCmdStoredProc
        Set ISqlOps_ExecuteCmd = oAdoCmd.Execute
        
    Case adCmdText, adCmdTable
        Select Case UCase$(Left$(oAdoCmd.CommandText, 1))
        Case "S"   '- it's a select stmt
            Set ISqlOps_ExecuteCmd = oAdoCmd.Execute
        Case Else  '- it's an update, insert, or delete stmt
            Call oAdoCmd.Execute(lRecs)
            ISqlOps_ExecuteCmd = lRecs
        End Select
        
    Case Else
        Set ISqlOps_ExecuteCmd = oAdoCmd.Execute
        
    End Select
End Function


Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    Select Case eiSeq
    Case esMAINSEQ:         SeqEnumToString = "MainSeq"
    Case esAPPT_ID:         SeqEnumToString = "APPT_ID"
    Case esAPPT_TRANS_ID:   SeqEnumToString = "APPT_TRANS_ID"
    Case esENC_TEMPLATE_ID: SeqEnumToString = "ENC_TEMPLATE_ID"
    Case esENCOUNTER_ID:    SeqEnumToString = "ENCOUNTER_ID"
    Case esIEN:             SeqEnumToString = "IEN"
    Case esPCM_CODE:        SeqEnumToString = "PCM_CODE"
    Case esPI_DATA_ID:      SeqEnumToString = "PI_DATA_ID"
    Case esUNIT_NUMBER:     SeqEnumToString = "UNIT_NUMBER"
    Case esUSER_NUMBER:     SeqEnumToString = "USER_NUMBER"
    Case esJDOE_NUMBER:     SeqEnumToString = "JDOE_NUMBER"
    Case esTEMP_SSN:        SeqEnumToString = "TEMP_SSN"
    End Select

End Function

Private Function ISqlOps_GetNextNcid() As Variant
'-- pulls the next value from the NCID sequence
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim lSeqNum As Long
    
    oRs.Open "Select Value from Sequences where seq_type = 'PROVIDER_NCID'", goITTOPSCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        lSeqNum = Val(oRs(0))
    End If
    goITTOPSCnn.Execute "update Sequences set value = " & lSeqNum + 1 & " where seq_type = 'PROVIDER_NCID'"
    ISqlOps_GetNextNcid = lSeqNum
    
    oRs.Close
    Set oRs = Nothing
End Function

Public Function ValidateProvider(ByRef ProviderNCID As String, _
                                 ByVal ProviderSSN As String, _
                                 ByVal ProviderFirstName As String, _
                                 ByVal ProviderMiddleName As String, _
                                 ByVal ProviderLastName As String, _
                                 ByVal CreateIfNotFound As Boolean, _
                        Optional ByVal FacilityNcid As String, _
                        Optional ByVal ClinicNCID As String) As Boolean

'called only by the data manager back end during decomposition of the XML message.
'Search for provider SSN in provider table.
' If found, update names (in provider and clinic-provider table) and return our ncid
' If not, create a new entry (local facility and clinic) and return our ncid
    
    '-- see if the provider is in the povider table
    
    Dim bNeedProvider As Boolean
    Dim bSuccess As Boolean
      
    On Error GoTo ErrHandler
'- see what's out there

    bNeedProvider = Not FindProvider(ProviderNCID, ProviderSSN)
   
    If Not bNeedProvider Then
'- we found the guy and have reset the providerncid (if needed) to our ncid.
        ValidateProvider = True
        Exit Function
    End If

'- so if we're here then:
'  we didn't find him.  should we create him?
    If Not CreateIfNotFound Then
        ValidateProvider = Not bNeedProvider
    Else
        DoEvents   '- settle down
        '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
        'ValidateProvider = BackfillProvider(ProviderNCID, ProviderSSN, ProviderFirstName, ProviderMiddleName, ProviderLastName)
        ValidateProvider = BackfillProvider(ProviderNCID, ProviderSSN, ProviderFirstName, ProviderMiddleName, ProviderLastName, FacilityNcid, ClinicNCID)
    End If

ErrHandler:

End Function

Public Function ValidateFacility(ByRef FacilityNcid, _
                                 ByRef FacilityName, _
                                 ByVal CreateIfNotFound As Boolean, _
                        Optional ByRef ClinicNCID, _
                        Optional ByRef ClinicName) As Boolean


    '-- see if the facility is in the facility table
    Dim bNeedFacility As Boolean
    Dim bNeedClinic As Boolean
    Dim bSuccess As Boolean
    
    On Error GoTo ErrHandler
'- see what's out there

    bNeedFacility = Not FindFacility(FacilityNcid)
'642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
'    bNeedClinic = Not FindClinic(FacilityNCID, ClinicNCID)
    
    If Not IsMissing(ClinicNCID) Then ''<SCR 24616 CC>
        bNeedClinic = Not FindClinic(FacilityNcid, ClinicNCID)
    End If

    If Not bNeedFacility And Not bNeedClinic Then
        ValidateFacility = True
        Exit Function
    End If
    
    If Not CreateIfNotFound Then
        ValidateFacility = Not bNeedFacility And Not bNeedClinic
        Exit Function
    End If
    
'- so we shouldn't create what we don't have
    bSuccess = True

    If bNeedFacility Then
        bSuccess = BackfillFacility(FacilityNcid, FacilityName)
    End If
    '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
    If Not IsMissing(ClinicNCID) Then '<SCR 24616 CC>
        If bNeedClinic Then
            bSuccess = bSuccess And BackfillClinic(FacilityNcid, ClinicNCID, ClinicName)
        End If
    End If
    
    ValidateFacility = bSuccess

ErrHandler:

End Function

Private Function FindFacility(ByVal FacilityNcid As Long) As Boolean
'_by ncid only
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    oRs.Open "Select facility_ncid from facility_ncid_correlation where facility_ncid = " & FacilityNcid, goITTOPSCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not oRs.EOF Then
        FindFacility = True
    Else
        FindFacility = False
    End If
    oRs.Close
    Set oRs = Nothing

End Function
Private Function FindClinic(ByVal FacilityNcid As Long, ByVal ClinicNCID As Long) As Boolean
'_ by ncid only


    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    If ClinicNCID = 0 Then
        FindClinic = True
        Exit Function
    End If
    
    oRs.Open "Select ncid from clinic where facility_ncid = " & FacilityNcid _
           & " and ncid = " & ClinicNCID, goITTOPSCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        FindClinic = True
    Else
        FindClinic = False
    End If
    
    oRs.Close
    Set oRs = Nothing

End Function

Private Function FindProvider(ByRef ProviderNCID As String, _
                              ByVal ProviderSSN As String) As Boolean

    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
'__ we must have an ssn
    If ProviderSSN = "" Then Exit Function
    
    oRs.Open "Select ncid from Provider where ssn = '" & ProviderSSN & "'", goITTOPSCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        ProviderNCID = "" & oRs("ncid")
        FindProvider = True
    Else
        FindProvider = False
    End If
    
    oRs.Close
    Set oRs = Nothing

End Function
'***********************************************************************
' changes: WRogers 4/3/2002 - refer to SCR 19356
' Facility_NCID_correlation table, type column is 10 characters in theater
' and cdr mode, we don't want to change it.
'
' added checks for ncids table before attempting to insert values
'***********************************************************************
Private Function BackfillFacility(ByVal FacilityNcid As String, ByVal FacilityName As String) As Boolean
    Dim rs As ADODB.Recordset

    If FacilityNcid = 0 Then Exit Function
    On Error GoTo ErrHandler
    
    If FacilityName = "" Or Len(FacilityName) > 10 Then FacilityName = "CHCS T ADT"
    
    goITTOPSCnn.Execute "insert into facility_ncid_correlation (facility_ncid,  type) values (" _
                & FacilityNcid & ",'" & FacilityName & "')"
   ' goITTOPSCnn.Execute "insert into facility_ncid_correlation (facility_ncid, context_ncid, type) values (" _
                & FacilityNcid & ",33962,'ADT')"
    
    Set rs = New ADODB.Recordset
    rs.Open "select ncid from ncids where ncid= " & FacilityNcid & " and domainncid =119", goITTOPSCnn
    If rs.EOF Then
        goITTOPSCnn.Execute "insert into ncids (ncid, domainncid, description) values (" _
            & FacilityNcid & ",119,'" & FacilityName & "')"
    End If
    
    rs.Close
    rs.Open "select ncid from ncids where ncid= " & FacilityNcid & " and domainncid =33962", goITTOPSCnn
    
    If rs.EOF Then
        goITTOPSCnn.Execute "insert into ncids (ncid, domainncid, description) values (" _
            & FacilityNcid & ",33962,'" & FacilityName & "')"
    End If
    rs.Close
    BackfillFacility = True
   
ErrHandler:
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function

Private Function BackfillClinic(ByVal FacilityNcid As String, _
                                ByVal ClinicNCID As String, _
                                ByVal ClinicName As String) As Boolean
    If ClinicNCID = 0 Then Exit Function
    
'_____ assumes the facility is already there  __________
    
    Dim meprs_code As String
    Dim meprs_ien As String
    On Error GoTo ErrHandler
    meprs_code = ISqlOps_GetNextID()
    meprs_ien = ClinicNCID
    
    If ClinicName = "" Or Len(ClinicName) > 40 Then ClinicName = "CHCSII-T ADT"
    
    goITTOPSCnn.Execute "insert into clinic (ncid, facility_ncid, location_type,meprs_code, meprs_ien, name) values (" _
                & ClinicNCID & "," & FacilityNcid & ",'CLINIC'," & meprs_code & "," & meprs_ien & ",'" & ClinicName & "')"
    
    goITTOPSCnn.Execute "insert into ncids (ncid, domainncid, description) values (" _
                & ClinicNCID & ",1003,'" & ClinicName & "')"

    goITTOPSCnn.Execute "insert into meprs_code(facility_ncid, ien, code, description) values (" _
                & FacilityNcid & "," & meprs_ien & "," & meprs_code & ",'" & ClinicName & "')"
    BackfillClinic = True
ErrHandler:
End Function

Private Function BackfillProvider(ByVal ProviderNCID As String, _
                                  ByVal ProviderSSN As String, _
                                  ByVal ProviderFirstName As String, _
                                  ByVal ProviderMiddleName As String, _
                                  ByVal ProviderLastName As String, _
                         Optional ByVal FacilityNcid As String, _
                         Optional ByVal ClinicNCID As String) As Boolean
     
    If Val(ProviderNCID) = 0 Then Exit Function
    
    Dim pcm_code    As String
    Dim meprs_ien   As String
    Dim sName       As String
    Dim Unit_Number As String
    Dim User_Number As String
    Dim IEN         As String
    
    Dim oRs         As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    pcm_code = ISqlOps_GetNextID(esPCM_CODE)
    IEN = ISqlOps_GetNextID(esIEN)
    Unit_Number = ISqlOps_GetNextID(esUNIT_NUMBER)
    User_Number = ISqlOps_GetNextID(esUSER_NUMBER)
    
    '- organize the name
    sName = UCase$(Trim$(ProviderLastName & ", " & ProviderFirstName & " " & ProviderMiddleName))
    If Trim$(sName) = "," Then
        sName = "(Unknown Provider)"
    End If
    
'__get the local facility and clinic ncids if needed
    '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
    If FacilityNcid = "" Or ClinicNCID = "" Then
    '
      If mFacilityNCID = "" Or mClinicNCID = "" Then
          oRs.Open "Select apkey, apvalue from osh_ini_settings " _
              & "where apsection = 'LOCAL_ENV'", _
              goITTOPSCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
          Do While Not oRs.EOF
              Select Case UCase$("" & oRs("apkey"))
              Case "FACILITY_NCID":       mFacilityNCID = oRs("apvalue")
              Case "CLINIC_NCID":         mClinicNCID = oRs("apvalue")
              End Select
              oRs.MoveNext
          Loop
          oRs.Close
          Set oRs = Nothing
      End If
    '<< Begin:SCR #25987;   Developer: Brian Mowbray 07/24/2002 03:25 PM
    ' Added this the "Else", otherwise these variables would always be empty if the Facility and Clinic was passed in
    Else
      mFacilityNCID = FacilityNcid & vbNullString
      mClinicNCID = ClinicNCID & vbNullString
    End If
    '>> End: SCR #25987;
    
    goITTOPSCnn.Execute "insert into hems_users (user_NCID, User_number,Unit_number) values ( " & ProviderNCID & "," & User_Number & "," & Unit_Number & ")"
    
    goITTOPSCnn.Execute "insert into provider (ncid, ien,pcm_CODE,fACILITYNCID, clinic_ncid,Unit_number, name) values (" _
                & ProviderNCID & "," & IEN & "," & pcm_code & "," & mFacilityNCID & "," & mClinicNCID & "," & Unit_Number & "," & SQLQuote(sName) & ")"
    
    
    BackfillProvider = True
ErrHandler:
  

End Function

Public Property Get ISqlOps_Value(Column As Variant) As Variant

    Dim index As Integer
    On Error GoTo ErrHandler
    If ISqlOps_BOF Or ISqlOps_EOF Then
        GoTo ErrHandler
    End If
    
    If IsNumeric(Column) Then
        index = Column
    Else
        index = ISqlOps_ColumnIndex(CStr(Column))
    End If
    
    ISqlOps_Value = mvarRecordSet.Fields(index).Value
'    If mvarRecordSet.Fields(index).Type = adDate Or mvarRecordSet.Fields(index).Type = adDBDate Or mvarRecordSet.Fields(index).Type = adDBTime Or mvarRecordSet.Fields(index).Type = adDBTimeStamp Then
'      If ISqlOps_Value <> #12/31/9999# Then
'          ' adjust to local from gmt
'          ISqlOps_Value = DateAdd("n", -lTimeBias, ISqlOps_Value)
'      End If
'    End If
    If IsNull(ISqlOps_Value) Then
        Select Case mvarRecordSet.Fields(index).Type
            Case adBoolean
                ISqlOps_Value = False
            Case adNumeric, adDecimal, adDouble, adSmallInt, adTinyInt, adCurrency, adSingle, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                ISqlOps_Value = 0
            Case Else
                ISqlOps_Value = ""
        End Select
    End If  ' IsNull(ISqlOps_Value)
ErrHandler:
End Property

Public Property Get ISqlOps_InfoType(Column As Variant) As String

    Dim index As Integer
    If (ISqlOps_BOF Or ISqlOps_EOF) Then
        GoTo ErrHandler
    End If
    
    If IsNumeric(Column) Then
        index = Column
    Else
        index = ISqlOps_ColumnIndex(CStr(Column))
    End If
    
    ISqlOps_InfoType = mvarRecordSet.Fields(index).Type  ' mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value.InfoType

ErrHandler:

End Property

Public Property Get ISqlOps_ColumnIndex(Name As String) As Integer

    Dim i As Integer
    Dim sText As String
    For i = 0 To ISqlOps_ColumnCount - 1
        If UCase(ISqlOps_ColumnName(i)) = UCase(Name) Then
            ISqlOps_ColumnIndex = i
            Exit Property
        End If
    Next i
    
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 0 To ISqlOps_ColumnCount - 1
        sText = sText & Chr(10) & mvarRecordSet.Fields(i).Name
    Next
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText
End Property

Public Property Get ISqlOps_ColumnName(index As Integer) As String

  
  ISqlOps_ColumnName = ""
    If Not (ISqlOps_BOF Or ISqlOps_EOF) Then
        If index > ISqlOps_ColumnCount Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & ISqlOps_ColumnCount & " colummns."
            Exit Property
        End If
        ISqlOps_ColumnName = mvarRecordSet.Fields(index).Name
    End If
End Property

Public Property Get ISqlOps_ColumnCount() As Integer

'    ColumnCount = mvarDSQLConnect.Table.Rows(mvarAbsolutePosition).Columns.Count
  ISqlOps_ColumnCount = mvarRecordSet.Fields.Count
End Property

Public Function ISqlOps_Execute(SqlStatement As String) As Variant
    ISqlOps_Execute = Execute(SqlStatement)
End Function
Public Function Execute(SqlStatement As String, Optional ByVal direct As Boolean = False) As Variant
        
    Dim lRecCnt     As Long
    Dim FirstInsPt  As Long
    Dim SecondInsPt As Long
    Dim ValuesStrPt As Long                 '*-- Coleman 040401
    Dim words()     As String
    Dim WhereClause As String
    
#If DEBUGON Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Execute", "Begin", DR, "SQL", SqlStatement
#End If
    ' SCR #24346;   Developer:  06/05/2002 11:05 AM
    ' Note: Function Changed:  Execute    Impact of Change:  added call to FixParens()
    
    '- ensure there are always spaces between parens and words, exclude quoted string
    SqlStatement = FixParans(SqlStatement)
    
    words = Split(SqlStatement, " ")
    mvarSQLStatement = SqlStatement

  If mvarRecordSet Is Nothing Then
    Set mvarRecordSet = New ADODB.Recordset
  Else
    If mvarRecordSet.State = adStateExecuting Then
      mvarRecordSet.Cancel
    End If
    
    If mvarRecordSet.State = adStateOpen Then
        On Error Resume Next
'        Debug.Print mvarRecordSet.EditMode
        mvarRecordSet.CancelUpdate
        mvarRecordSet.Close
        On Error GoTo 0
    End If
    
  End If
  
  If goITTOPSCnn Is Nothing Then
'-- ITT
    Call OpenDB
     
  End If
  
  If goITTOPSCnn.State = 0 Then
    goITTOPSCnn.Open
  End If
  
  If UBound(words) > 1 Then
    Select Case UCase$(Left$(SqlStatement, 1))
    
    Case "D"
        '- is as "Delete from tablename where field...
        If Not direct And Not DoNoAuditing(words(2)) Then
            If InStr(LCase$(SqlStatement), " where ") Then
                WhereClause = Mid$(SqlStatement, InStr(LCase$(SqlStatement), " where "))
            Else
                WhereClause = ""
            End If
            goITTOPSCnn.Execute "update " & words(2) & " set UpdatedBy = " & SQLQuote(UserNCID) _
                        & ", UpdatedOn = " & SQLDate(Now()) & WhereClause, lRecCnt
        End If
        goITTOPSCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "I"    '- may be "Insert into Tablename ( field1....) values (value1...)"
                '- or     "Insert into TableName (field1...) select ....."
                
        If InStr(words(2), "(") > 1 Then words(2) = Left$(words(2), InStr(words(2), "(") - 1)
        
        If Not direct And Not DoNoAuditing(words(2)) Then
            FirstInsPt = InStr(SqlStatement, "(")
            '*-- Coleman 040401: Handles case where no space before parentheses - "values("
            '*-- If InStr(LCase(SqlStatement), " values ") > 0 Then
            '*--    SecondInsPt = InStr(FirstInsPt + 1, SqlStatement, "(")
            ValuesStrPt = InStr(LCase(SqlStatement), " values")         '*-- Coleman 040401
            If ValuesStrPt > 0 Then                                     '*-- Coleman 040401
                SecondInsPt = InStr(ValuesStrPt, SqlStatement, "(")     '*-- Coleman 040401
            Else
                SecondInsPt = InStr(LCase(SqlStatement), " select ") + 7
            End If
       
            SqlStatement = Stuff(SqlStatement, SecondInsPt, SQLQuote(UserNCID) & ", " & SQLDate(Now()) & ", ")
            SqlStatement = Stuff(SqlStatement, FirstInsPt, "CreatedBy, CreatedOn, ")
        End If
        
        goITTOPSCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "U"    '- must be "Update TableName Set Field1 = ..."
    '- get the table name
       
        If Not direct And Not DoNoAuditing(words(1)) Then
            FirstInsPt = InStr(UCase$(SqlStatement), "SET")
            SqlStatement = Stuff(SqlStatement, FirstInsPt + 3, " UpdatedBy = " & SQLQuote(UserNCID) & ", UpdatedOn = " & SQLDate(Now()) & ", ")
        End If
        goITTOPSCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "S"
       mvarRecordSet.Open SqlStatement, goITTOPSCnn, adOpenDynamic, adLockPessimistic, adCmdText
       'Set ISqlOps_Execute = mvarRecordSet
    Case "E"
        goITTOPSCnn.Execute SqlStatement, lRecCnt, adCmdText
    End Select
  Else  '- ubound(words) > 1
     mvarRecordSet.Open SqlStatement, goITTOPSCnn, adOpenKeyset, adLockOptimistic, adCmdTable
  End If
  
#If DEBUGON Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Execute", "End", DR
#End If
    
  
End Function
Public Sub ISqlOps_ExecuteMulti(SqlStatements As Collection)
    Call ExecuteMulti(SqlStatements)
End Sub

Public Sub ExecuteMulti(SqlStatements As Collection, Optional ByVal direct As Boolean = False)
  Dim vStmt As Variant
  On Error GoTo ErrHandler
  
  goITTOPSCnn.BeginTrans
  
  For Each vStmt In SqlStatements
    mvarSQLStatement = vStmt
    Execute CStr(vStmt), direct
  Next vStmt
      
  goITTOPSCnn.CommitTrans
  Exit Sub
  
ErrHandler:
  goITTOPSCnn.RollbackTrans
  
End Sub


Public Sub ISqlOps_MovePrevious()
    If mvarRecordSet.State = 0 Or mvarRecordSet.BOF Then Exit Sub
    mvarRecordSet.MovePrevious
End Sub

Public Sub ISqlOps_MoveNext()
    If mvarRecordSet.State = 0 Or mvarRecordSet.EOF Then Exit Sub
    mvarRecordSet.MoveNext
End Sub

Public Sub ISqlOps_MoveLast()
    If mvarRecordSet.State = 0 Then Exit Sub
    mvarRecordSet.MoveLast
    
End Sub

Public Sub ISqlOps_MoveFirst()
    If mvarRecordSet.State = 0 Then Exit Sub
    mvarRecordSet.MoveFirst

End Sub

Public Property Get ISqlOps_RecordCount() As Long
    On Error Resume Next
    ISqlOps_RecordCount = mvarRecordSet.RecordCount
End Property

Public Property Get ISqlOps_Source() As String:  ISqlOps_Source = mvarSQLStatement: End Property
Public Property Get ISqlOps_EOF() As Boolean:    ISqlOps_EOF = mvarRecordSet.EOF:   End Property
Public Property Get ISqlOps_BOF() As Boolean:    ISqlOps_BOF = mvarRecordSet.BOF:   End Property

Private Sub Class_Initialize()
#If DEBUGON Then
  Set DL = GetObject("", "dlog.cdebuglog")
  If Not DL Is Nothing Then DL.debugStart
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "", DR
#End If
    
    GetTimeZoneInformation TZInfo
    lTimeBias = TZInfo.Bias

    Set mobjShared = New CWShared
    UserNCID = mobjShared.CurrentUserNCID

    glITTOPSCnnCnt = glITTOPSCnnCnt + 1
    If goITTOPSCnn Is Nothing Then
        Set goITTOPSCnn = New ADODB.Connection
        goITTOPSCnn.ConnectionString = GetConnectionString
        goITTOPSCnn.ConnectionTimeout = 30
        goITTOPSCnn.CursorLocation = adUseClient
        If goITTOPSCnn.State = 0 Then goITTOPSCnn.Open
'        If goITTOPSCnn.State = 1 Then
            'THIS JUST EXITS IF USING SQL SERVER
'            Call SetDBDateFormat(goITTOPSCnn)
'        End If
    End If

  Set mvarRecordSet = New ADODB.Recordset

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
'- handle the recordset
    If mvarRecordSet.State > 0 Then
        If Not (mvarRecordSet.EOF Or mvarRecordSet.BOF) Then
            mvarRecordSet.CancelUpdate      '- this will error if not in an update, but we're closing...
        End If
        mvarRecordSet.Close
    End If
    Set mvarRecordSet = Nothing
    
'- handle the connection
    glITTOPSCnnCnt = glITTOPSCnnCnt - 1
    If glITTOPSCnnCnt <= 0 Then
        If goITTOPSCnn.State > 0 Then goITTOPSCnn.Close
        Set goITTOPSCnn = Nothing
    End If
#If DEBUGON Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "", DR
#End If
    
End Sub

Private Function OpenDB()
  'Set goITTOPSCnn = DE.Cnn

    Set goITTOPSCnn = New ADODB.Connection
    
    goITTOPSCnn.ConnectionString = GetConnectionString
    '  goITTOPSCnn.ConnectionTimeout = 30
    goITTOPSCnn.Open
'    If goITTOPSCnn.State = 1 And goITTOPSCnn.Properties("DBMS Name") = "Microsoft SQL Server" Then    '-- ITT
        'THIS DOESN'T MAKE SENSE BECAUSE SetDBDateFormat EXITS IF IN SQL SERVER MODE
'        Call SetDBDateFormat(goITTOPSCnn)
'    End If
  
End Function

Private Function DoNoAuditing(sTable As String) As Boolean
    Dim sMember As Variant
    
    If mNoAudit Is Nothing Then LoadNoAudit
    
    For Each sMember In mNoAudit
        If LCase(sMember) = LCase(sTable) Then
            DoNoAuditing = True
            Exit For
        End If
    Next sMember
End Function

Private Sub LoadNoAudit()
    Set mNoAudit = New Collection
    
    Dim oRs As ADODB.Recordset
    Dim oSql As ICHCSII_SQL.ISqlOps
    Set oSql = New ITT_SQLOPS
    
    oSql.Execute "DoNotAudit"
    Do While Not oSql.EOF
        mNoAudit.Add "" & oSql(0)
        oSql.MoveNext
    Loop
    
    Set oSql = Nothing

End Sub


'<< These are pass-through functions , same as isqlopsex_
Private Property Get ISQLDirect_BOF() As Boolean: ISQLDirect_BOF = ISqlOps_BOF: End Property
Private Property Get ISQLDirect_ColumnCount() As Integer: ISQLDirect_ColumnCount = ISqlOps_ColumnCount: End Property
Private Property Get ISQLDirect_ColumnIndex(Name As String) As Integer: ISQLDirect_ColumnIndex = ISqlOps_ColumnIndex(Name): End Property
Private Property Get ISQLDirect_ColumnName(index As Integer) As String: ISQLDirect_ColumnName = ISqlOps_ColumnName(index): End Property
Private Property Get ISQLDirect_EOF() As Boolean: ISQLDirect_EOF = ISqlOps_EOF: End Property
Private Function ISQLDirect_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.edSeq = edMAINSEQ) As Variant: ISQLDirect_GetNextID = ISqlOps_GetNextID(eSeqName): End Function
Private Function ISQLDirect_GetNextNcid() As Variant: ISQLDirect_GetNextNcid = ISqlOps_GetNextNcid: End Function
Private Property Get ISQLDirect_InfoType(Column As Variant) As String: ISQLDirect_InfoType = ISqlOps_InfoType(Column): End Property
Private Sub ISQLDirect_MoveFirst(): ISqlOps_MoveFirst: End Sub
Private Sub ISQLDirect_MoveLast(): ISqlOps_MoveLast: End Sub
Private Sub ISQLDirect_MoveNext(): ISqlOps_MoveNext: End Sub
Private Sub ISQLDirect_MovePrevious(): ISqlOps_MovePrevious: End Sub
Private Property Get ISQLDirect_RecordCount() As Long: ISQLDirect_RecordCount = ISqlOps_RecordCount: End Property
Private Property Get ISQLDirect_Source() As String: ISQLDirect_Source = ISqlOps_Source: End Property
Private Function ISQLDirect_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant: ISQLDirect_ExecuteCmd = ISqlOps_ExecuteCmd(oAdoCmd): End Function
Private Property Get ISQLDirect_Value(Column As Variant) As Variant: ISQLDirect_Value = ISqlOps_Value(Column): End Property
Private Function ISQLDirect_ValidateFacility(ByVal FacilityNcid As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal ClinicNCID As Long = 0&) As Boolean
    ISQLDirect_ValidateFacility = ValidateFacility(FacilityNcid, CreateIfNotFound, ClinicNCID)
End Function
'>> end of pass-through
Private Function ISQLDirect_Execute(SqlStatement As String) As Variant
    ISQLDirect_Execute = Execute(SqlStatement, True)
End Function
Private Sub ISQLDirect_ExecuteMulti(SqlStatements As Collection)
    ExecuteMulti SqlStatements, True
End Sub



'<< these are all 'pass-throughs' to the isqlops_
Private Property Get ISqlOpsEx_BOF() As Boolean: ISqlOpsEx_BOF = ISqlOps_BOF: End Property
Private Property Get ISqlOpsEx_ColumnCount() As Integer: ISqlOpsEx_ColumnCount = ISqlOps_ColumnCount: End Property
Private Property Get ISqlOpsEx_ColumnIndex(Name As String) As Integer: ISqlOpsEx_ColumnIndex = ISqlOps_ColumnIndex(Name): End Property
Private Property Get ISqlOpsEx_ColumnName(index As Integer) As String: ISqlOpsEx_ColumnName = ISqlOps_ColumnName(index): End Property
Private Property Get ISqlOpsEx_EOF() As Boolean: ISqlOpsEx_EOF = ISqlOps_EOF: End Property
Private Function ISqlOpsEx_Execute(SqlStatement As String) As Variant: ISqlOpsEx_Execute = ISqlOps_Execute(SqlStatement): End Function
Private Sub ISqlOpsEx_ExecuteMulti(SqlStatements As Collection): ExecuteMulti SqlStatements: End Sub
Private Function ISqlOpsEx_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.exSeq = exMAINSEQ) As Variant: ISqlOpsEx_GetNextID = ISqlOps_GetNextID(eSeqName): End Function
Private Function ISqlOpsEx_GetNextNcid() As Variant: ISqlOpsEx_GetNextNcid = ISqlOps_GetNextNcid: End Function
Private Property Get ISqlOpsEx_InfoType(Column As Variant) As String: ISqlOpsEx_InfoType = ISqlOps_InfoType(Column): End Property
Private Sub ISqlOpsEx_MoveFirst(): ISqlOps_MoveFirst: End Sub
Private Sub ISqlOpsEx_MoveLast(): ISqlOps_MoveLast: End Sub
Private Sub ISqlOpsEx_MoveNext(): ISqlOps_MoveNext: End Sub
Private Sub ISqlOpsEx_MovePrevious(): ISqlOps_MovePrevious: End Sub
Private Property Get ISqlOpsEx_RecordCount() As Long: ISqlOpsEx_RecordCount = ISqlOps_RecordCount: End Property
Private Property Get ISqlOpsEx_Source() As String: ISqlOpsEx_Source = ISqlOps_Source: End Property
Private Function ISqlOpsEx_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant: ISqlOpsEx_ExecuteCmd = ISqlOps_ExecuteCmd(oAdoCmd): End Function

Private Function ISqlOpsEx_SQLDate(ByVal dDate As Date) As String
    ISqlOpsEx_SQLDate = SQLDate(Date)
End Function

Private Function ISqlOpsEx_SQLQuote(ByVal sText As String) As String
    ISqlOpsEx_SQLQuote = SQLQuote(sText)
End Function

Private Function ISqlOpsEx_ValidateFacility(ByVal FacilityNcid As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal ClinicNCID As Long = 0&) As Boolean
    '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
  ISqlOpsEx_ValidateFacility = ValidateFacility(FacilityNcid, "CHCSII-T ADT", CreateIfNotFound, ClinicNCID, "CHCSII-T ADT")
End Function
' SCR #439;   Developer:  06/05/2002 04:17 PM
' Note: Function Changed:  ValidateFacility2  ValidateProvider2  SQLDate  SQLCase  SQLQuote    Impact of Change:

Public Function ISqlOpsEx_ValidateFacility2(ByRef FacilityNcid, _
                                           ByRef FacilityName, _
                                           ByVal CreateIfNotFound As Boolean, _
                                  Optional ByRef ClinicNCID, _
                                  Optional ByRef ClinicName) As Boolean

    ISqlOpsEx_ValidateFacility2 = ValidateFacility(FacilityNcid, FacilityName, CreateIfNotFound, ClinicNCID, ClinicName)
End Function
                                
Private Function ISqlOpsEx_ValidateProvider(ByVal UserNCID As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal FacilityNcid As Long = 0&, Optional ByVal ClinicNCID As Long = 0&) As Boolean
  '642 Patch 2 changes made to synchronize baseline. Brian Mowbray 7/9/2002
  ISqlOpsEx_ValidateProvider = ValidateProvider(CStr(UserNCID), "", "", "", "", CreateIfNotFound, FacilityNcid, ClinicNCID)
End Function

Public Function ISqlOpsEx_ValidateProvider2(ByRef ProviderNCID As String, _
                                           ByVal ProviderSSN As String, _
                                           ByVal ProviderFirstName As String, _
                                           ByVal ProviderMiddleName As String, _
                                           ByVal ProviderLastName As String, _
                                           ByVal CreateIfNotFound As Boolean) As Boolean
    ISqlOpsEx_ValidateProvider2 = ValidateProvider(ProviderNCID, ProviderSSN, ProviderFirstName, ProviderMiddleName, ProviderLastName, CreateIfNotFound)
End Function


Private Property Get ISqlOpsEx_Value(Column As Variant) As Variant: ISqlOpsEx_Value = ISqlOps_Value(Column): End Property

Private Function ISqlOpsex_AddNew() As Boolean
   If mvarRecordSet Is Nothing Then Exit Function
   On Error GoTo ErrHandler
   
   mvarRecordSet.AddNew
   ISqlOpsex_AddNew = True

ErrHandler:
   
End Function

Private Function ISqlOpsEx_Update() As Boolean
    On Error GoTo ErrHandler
    mvarRecordSet.Update
    ISqlOpsEx_Update = True
ErrHandler:
End Function

Private Function ISqlOpsEx_Cancel() As Boolean
    On Error GoTo ErrHandler
    mvarRecordSet.Cancel
    ISqlOpsEx_Cancel = True
ErrHandler:
End Function

Private Property Let ISqlOpsEx_Value(Column As Variant, ByVal RHS As Variant)
    Dim index As Integer
    On Error GoTo ErrHandler
    
    If ISqlOps_BOF Or ISqlOps_EOF Then
        GoTo ErrHandler
    End If
        
'- get the column
    If IsNumeric(Column) Then
        index = Column
    Else
        index = ISqlOps_ColumnIndex(CStr(Column))
    End If
    
    If mvarRecordSet.Fields(index).Type = adDate Or mvarRecordSet.Fields(index).Type = adDBDate Or mvarRecordSet.Fields(index).Type = adDBTime Or mvarRecordSet.Fields(index).Type = adDBTimeStamp Then
'      If RHS <> #12/31/9999# Then
'          '- adjust from local to Zulu time by subtracting the time bias
'          RHS = AdjustDate(RHS)
'      End If
    End If

    mvarRecordSet.Fields(index).Value = RHS

ErrHandler:

End Property



Private Function GetConnectionString() As String

    'SF 2-2-04 Use GetSQLServerConnectionString when in ITT mode.
    
    GetConnectionString = SQLSVR_Conn
    GetConnectionString = GetSQLServerConnectionString(GetConnectionString)

End Function

Private Function ISqlOpsex_ExecuteSP(ByVal SPName As String, _
                                         ParamArray aParams() As Variant) As Variant: ISqlOpsex_ExecuteSP = False: End Function





