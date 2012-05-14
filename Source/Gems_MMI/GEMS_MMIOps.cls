VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_MMIOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'--DB Connection
Private sql          As String
Private mobjDAS      As ICHCSII_DAS

'--Patient Object
Private mcolPatients As Collection     '- a coll of gems-patient
Private mobjCurrPat  As GEMS_MMI       '- the current or active patient
Private mvarEOF      As Boolean        '- curr pat = last pat in coll
Private mvarPatIdx   As Long           '- the current patient's index in the pat coll

Private mobjShared   As CWShared       '*-- used to display error dialog

Public Enum Search_GenericID
    SRCH_FirstLetterLastName_LastFourPatientSSN
    SRCH_FirstLetterLastName_LastFourSponsorSSN
    SRCH_FMP_SponsorSSN
    SRCH_FMP
    SRCH_SponsorSSN
    SRCH_ID_None
End Enum

'added WRogers 8/24/2001
Public Enum Search_GenericData
    SRCH_Generic_None
    SRCH_PERSONAL_RELIABILTY_PROGRAM = 204689
    SRCH_FLYING_STATUS = 217201
    SRCH_JUMPING_STATUS = 217202
    SRCH_DIVING_STATUS = 217203
    SRCH_ON_MOBILITY = 217204
    SRCH_MILITARY_POLICE = 217205
    SRCH_SUBMARINE = 217206
    SRCH_UIC = 14501722
    SRCH_ENROLLMENT_FACILITY = 14501723
    SRCH_COMMAND_SECURITY = 204687
    SRCH_COMMAND_INTEREST = 204685
    SRCH_FMP_SPONSOR_SSN = 14511309
    SRCH_FMP = 204708
    SRCH_CHAMPUS_RANK_GRADE = 1015
    SRCH_PATIENT_CATEGORY = 14501721
    SRCH_ACV_ENROLLMENT_STATUS = 204696
    SRCH_RECORDS_MAINTAINED_AT = 14504166
End Enum

Public Enum Search_PrimaryField
    SRCH_LastName
    SRCH_GenericID
    SRCH_GenericData
    SRCH_SSN
    SRCH_None
End Enum

Public Property Let Category(RHV As String)
'SCRs 19698, 19700 and 19701 - 2/19/02 - Walker - Used In Gems
'Demographics references this as the Patient Category value
    mobjCurrPat.Category = RHV
End Property

Public Property Get CurrentPatient() As GEMS_MMI
  Set CurrentPatient = mobjCurrPat
End Property
Public Property Set CurrentPatient(RHV As GEMS_MMI)
  Set mobjCurrPat = RHV
End Property
'Private mobjPatCrit As GEMS_MMICriteria


'WRogers 8/24/2001
' added parms for UIC search in the MMI_GenericData
Public Function DoSearch(menuGenericSearchType As Search_GenericID, _
                         menuPrimarySearchType As Search_PrimaryField, _
                         mvarFirstLetterLastName As String, _
                         mvarLastFourPatientSSN As String, _
                         mvarLastFourSponsorSSN As String, _
                         mvarLastName As String, _
                         mvarFirstName As String, _
                         mvarDOB As String, _
                         mvarSSN As String, _
                         mvarFMP As String, _
                         mvarSponsorSSN As String, _
                         mvarSEX As String, _
                         mvarNameType As String, _
                         menuGenericDataType As Search_GenericData, _
                         mvarGenericData As String) As Collection

    Dim NameWhereClause     As String
    Dim GenIDWhereClause    As String
    Dim IDWhereClause       As String
    Dim GenDataWhereClause  As String
    Dim bUseName            As Boolean
    Dim bUseGenericID       As Boolean
    Dim bUseID              As Boolean
    Dim bUseGenericData     As Boolean
    Dim sNameStr            As String       '*-- Coleman 111401: SCR 18601
    Dim oDatabase As CHCSII_CONN.Conn
    Dim rst As Adodb.Recordset
    
    Set oDatabase = New CHCSII_CONN.Conn

    If menuPrimarySearchType = SRCH_LastName Then
        '*-- Coleman 111401: SCR 18601
        sNameStr = SQLQuote(UCase(mvarLastName))
        sNameStr = Left$(sNameStr, Len(sNameStr) - 1) & "%'"
        NameWhereClause = NameWhereClause & " and Last_Name like " & sNameStr
        bUseName = True
    End If

    If mvarFirstName <> "" Then
        '*-- Coleman 111401: SCR 18601
        sNameStr = SQLQuote(UCase(mvarFirstName))
        sNameStr = Left$(sNameStr, Len(sNameStr) - 1) & "%'"
        NameWhereClause = NameWhereClause & " and First_Name like " & sNameStr
        bUseName = True
    End If

    If NameWhereClause <> "" Then
      NameWhereClause = Mid$(NameWhereClause, 5) & " and Name_Type_cid = " & NCID_CURRENT_NAME
    End If

    If mvarDOB <> "" Then
        '*-- Coleman 062001: Fix Oracle data format bug and change DOB search
        '*-- to a range like MMI object (+/- 2 years).
        'IDWhereClause = IDWhereClause & " and Birth_date = '" & mvarDOB & "'"
        IDWhereClause = IDWhereClause & " and (Birth_date >= " & oDatabase.SQLDate(DateAdd("yyyy", -2, mvarDOB))
        IDWhereClause = IDWhereClause & " and Birth_date <= " & oDatabase.SQLDate(DateAdd("yyyy", 2, mvarDOB)) & ")"
        bUseID = True
    End If

    '*-- Coleman 060701: SCR 13341
    If (mvarSSN <> "") And (menuGenericSearchType <> Search_GenericID.SRCH_FirstLetterLastName_LastFourPatientSSN) Then
        IDWhereClause = IDWhereClause & " and SSN = '" & mvarSSN & "'"
        bUseID = True
    End If

    If mvarSEX <> "" Then
        IDWhereClause = IDWhereClause & " and Sex = '" & mvarSEX & "'"
        bUseID = True
    End If

    If IDWhereClause <> "" Then
      IDWhereClause = Mid$(IDWhereClause, 5)
    End If

    If menuPrimarySearchType = SRCH_GenericID Then
        Select Case menuGenericSearchType
            Case Search_GenericID.SRCH_FirstLetterLastName_LastFourPatientSSN
               GenIDWhereClause = GenIDWhereClause & " and id_type_ncid = " & NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_PATIENT_SSN _
                  & " and id_value ='" & mvarFirstLetterLastName & mvarLastFourPatientSSN & "'"
               bUseGenericID = True
            Case Search_GenericID.SRCH_FirstLetterLastName_LastFourSponsorSSN
               GenIDWhereClause = GenIDWhereClause & " and id_type_ncid = " & NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_SPONSOR_SSN _
                  & " and id_value ='" & mvarFirstLetterLastName & mvarLastFourSponsorSSN & "'"
               bUseGenericID = True

            Case Search_GenericID.SRCH_FMP_SponsorSSN
                'wrong NCID and missing FMP value SCR 19044
               GenIDWhereClause = GenIDWhereClause & " and id_type_ncid = " & NCID_FMP_SPONSOR_SSN _
                  & " and id_value ='" & Format(mvarFMP & mvarSponsorSSN, "00-000-00-0000") & "'"
               bUseGenericID = True
            '*-- Coleman 060701: SCR 13334
            Case Search_GenericID.SRCH_SponsorSSN
               GenIDWhereClause = GenIDWhereClause & " and id_type_ncid = " & NCID_SPONSOR_SSN & " and id_value ='" & mvarSponsorSSN & "'"
               bUseGenericID = True
               
        End Select
    Else
        'Check FMP or Sponsor SSN
        Select Case menuGenericSearchType
            Case Search_GenericID.SRCH_SponsorSSN
               GenIDWhereClause = GenIDWhereClause & " and id_type_ncid = " & NCID_SPONSOR_SSN _
                  & " and id_value ='" & mvarSponsorSSN & "'"
               bUseGenericID = True
            Case Search_GenericID.SRCH_FMP
               GenIDWhereClause = GenIDWhereClause & " and id_type_ncid = " & NCID_FMP _
                  & " and id_value ='" & mvarFMP & "'"
               bUseGenericID = True

        End Select
    End If
    
    'added WRogers 8/24/2001
    If menuGenericDataType Then
        GenDataWhereClause = " type_ncid =" & menuGenericDataType & _
                            " and value_text ='" & mvarGenericData & "' "
        bUseGenericData = True
    End If

    If GenIDWhereClause <> "" Then
      GenIDWhereClause = Mid$(GenIDWhereClause, 5)
    End If

    
    '---execute the patient search---

Dim SelectClause As String
Dim JoinClause   As String
   JoinClause = ""
   SelectClause = ""
  
  
  If bUseName And bUseID And bUseGenericID And bUseGenericData Then
        SelectClause = "select distinct mmi_id.unit_number from mmi_id,mmi_name,mmi_generic_id, mmi_genericdata "
        JoinClause = " and mmi_id.unit_number(+) =  mmi_name.unit_number and " _
        & "mmi_id.unit_number(+) = mmi_generic_id.unit_number and " _
        & "mmi_id.unit_number = mmi_genericdata.unit_number"

  ElseIf bUseName And bUseID And bUseGenericID Then
        SelectClause = "select distinct mmi_id.unit_number from mmi_id, mmi_generic_id, mmi_name "
        JoinClause = " and mmi_id.unit_number = mmi_generic_id.unit_number" _
            & " and mmi_id.unit_number(+) = mmi_name.unit_number"

  ElseIf bUseID And bUseGenericID And bUseGenericData Then
        SelectClause = "select distinct mmi_id.unit_number from mmi_id, mmi_generic_id, mmi_genericdata "
        JoinClause = " and mmi_id.unit_number(+) = mmi_generic_id.unit_number" _
            & " and mmi_id.unit_number = mmi_genericdata.unit_number"

  ElseIf bUseName And bUseGenericID And bUseGenericData Then
        SelectClause = "select distinct mmi_name.unit_number from mmi_name, mmi_generic_id, mmi_genericdata "
        JoinClause = " and mmi_name.unit_number(+) = mmi_generic_id.unit_number" _
            & " and mmi_name.unit_number = mmi_genericdata.unit_number"

  ElseIf bUseName And bUseID And bUseGenericData Then
        SelectClause = "select distinct mmi_name.unit_number from mmi_name, mmi_id, mmi_genericdata "
        JoinClause = " and mmi_name.unit_number(+) = mmi_id.unit_number" _
            & " and mmi_name.unit_number = mmi_genericdata.unit_number"

  ElseIf bUseName And bUseGenericData Then
        SelectClause = "select distinct mmi_name.unit_number from mmi_name, mmi_genericdata "
        JoinClause = " and mmi_name.unit_number(+) = mmi_genericdata.unit_number"

  ElseIf bUseName And bUseID Then
        SelectClause = "select distinct mmi_id.unit_number from mmi_id, mmi_name "
        JoinClause = " and mmi_id.unit_number(+) =  mmi_name.unit_number"

  ElseIf bUseName And bUseGenericID Then
        SelectClause = "select distinct mmi_name.unit_number from mmi_name, mmi_generic_id "
        JoinClause = " and  mmi_name.unit_number (+) = mmi_generic_id.unit_number"

  ElseIf bUseID And bUseGenericID Then
        SelectClause = "select distinct mmi_id.unit_number from mmi_id, mmi_generic_id "
        JoinClause = " and mmi_id.unit_number(+) = mmi_generic_id.unit_number"

  ElseIf bUseGenericID And bUseGenericData Then
    SelectClause = "select distinct mmi_generic_id.unit_number from mmi_generic_id, mmi_genericdata "
    JoinClause = " and mmi_generic_id.unit_number(+) = mmi_genericdata.unit_number"

  ElseIf bUseID And bUseGenericData Then
       SelectClause = "select distinct mmi_id.unit_number from mmi_id, mmi_genericdata "
        JoinClause = " and mmi_id.unit_number(+) = mmi_genericdata.unit_number"

  ElseIf bUseName Then
   SelectClause = "select distinct mmi_name.unit_number from mmi_name"

  ElseIf bUseID Then
   SelectClause = "select distinct mmi_id.unit_number from mmi_id"

  ElseIf bUseGenericID Then
   SelectClause = "select distinct mmi_generic_id.unit_number from mmi_generic_id"

  ElseIf bUseGenericData Then
   SelectClause = "select distinct mmi_genericdata.unit_number from mmi_genericdata"

  Else
    '--error
  End If
  
    Dim sTemp As String
    
    sql = SelectClause & vbCrLf & " WHERE " & NameWhereClause
    
    If IDWhereClause <> "" Then
        sTemp = " and " & IDWhereClause & vbCrLf
    End If
    If GenIDWhereClause <> "" Then
       sTemp = sTemp & " and " & GenIDWhereClause & vbCrLf
    End If
    If GenDataWhereClause <> "" Then
        sTemp = sTemp & " and " & GenDataWhereClause & vbCrLf
    End If
    
    If JoinClause <> "" Then
       JoinClause = Mid$(JoinClause, 5)
       sTemp = sTemp & " and " & JoinClause & vbCrLf
    End If
    
    If sTemp <> "" Then
        If NameWhereClause = "" Then
            sTemp = Mid$(sTemp, 5)
        End If
    End If
    
    sql = sql & sTemp
 
  Set rst = mobjDAS.OpenRecordset(sql)
Dim oPatient As GEMS_MMI
  
  Set mcolPatients = New Collection
  
  Do While Not rst.EOF
    '*-- To Do: Define a GetPatientLight function that returns only a small
    '*-- subset of the patient info...
    Set oPatient = GetPatient(rst(0))
    If Not oPatient Is Nothing Then
        LoadSponsor oPatient, oPatient.SponsorSSN
        mcolPatients.Add oPatient, CStr(oPatient.NCID)
    Else
        mobjShared.ShowVBError 0, "Error loading patient.", "GEMS_MMIOps.DoSearch", "Patient Error", vbExclamation
    End If
    rst.MoveNext
  Loop
  rst.Close

  If mcolPatients.Count > 0 Then
    mvarPatIdx = 1
    Set mobjCurrPat = mcolPatients(mvarPatIdx)
    mvarEOF = False
  Else
    mvarPatIdx = 0
    mvarEOF = True
  End If
  
  Set DoSearch = mcolPatients
  
  Set oPatient = Nothing
  If rst.State > 0 Then rst.Close
  Set rst = Nothing
  
 End Function

Public Function MoveNext() As Boolean

  mvarPatIdx = mvarPatIdx + 1
  If mvarPatIdx <= mcolPatients.Count Then
    Set mobjCurrPat = mcolPatients(mvarPatIdx)
    mvarEOF = False
    MoveNext = True
  Else
    mvarPatIdx = 0
    mvarEOF = True
    MoveNext = False
  End If
    
End Function

Public Function MoveFirst() As Boolean

  If mcolPatients.Count > 0 Then
    mvarPatIdx = 1
    Set mobjCurrPat = mcolPatients(mvarPatIdx)
    mvarEOF = False
    MoveFirst = True
  Else
    mvarPatIdx = 0
    mvarEOF = True
    MoveFirst = False
  End If

End Function

Public Function ClearPatient() As Boolean
  Set mcolPatients = New Collection
  mvarPatIdx = 0
  mvarEOF = True
  
End Function

Public Function GetUnitNumberFromSSN(sSSN As String, Optional bPatientSearch As Boolean = False, Optional lRecs As Long) As String
'2/21/2002 - Trackers 17993 and 19414 - Added optional parms:
'bPatientSearch = to indicate caller is Patient Search
'lRecs = to hold number of recs found
    
    Dim rs As Adodb.Recordset
    Dim sSQL As String
    
    'MAG Tracker #40438 08/22/2003
    sSQL = "SELECT unit_number FROM mmi_id" _
        & " WHERE ssn = '" & sSSN & "'"
    'End Tracker #40438
    
Get_Row:
    Set rs = mobjDAS.OpenRecordset(sSQL)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            '*-- Coleman 092701: May want to handle multiple records for one
            '*-- SSN differently, but for now at least flag the error.
            rs.MoveFirst
            
            'Trackers 17993 and 19414 - 2/21/02 - Walker - Return number of recs found
            lRecs = rs.RecordCount
            
            If rs.RecordCount > 1 Then
            
        '1/4/2002 - Trackers 17993 and 19414 - Suppress duplicate error message
        '1/4/2002        mobjShared.ShowVBError 0, _
                        "Database error: multiple patients for SSN " & SSN, _
                        "GetUnitNumberFromSSN", "GEMS_MMIOps", vbExclamation
                GetUnitNumberFromSSN = ""
                
            Else
                GetUnitNumberFromSSN = rs("unit_number")

            End If
        End If
        rs.Close
    End If
    
End Function

Public Function LoadPatient(NCID As Long, MedRecNum As String) As Boolean
  
  Dim sql As String
  Dim WhereClause As String
  Dim oPx As GEMS_MMI
  Dim rs As Adodb.Recordset
  
  Debug.Print "Begin LoadPatient: Unit_Number=" & NCID
  
  If Not NCID = 0 Then
    WhereClause = WhereClause & " and unit_number =" & NCID
  End If
  
'  If Not MedRecNum = "" Then
'    WhereClause = WhereClause & " and govt_health_number ='" & MedRecNum & "'"
'  End If
  
  If WhereClause = "" Then Exit Function
  
  WhereClause = Mid$(WhereClause, 5)
  
  sql = "SELECT unit_number" _
    & " FROM mmi_id" _
    & " WHERE " & WhereClause
  Set rs = mobjDAS.OpenRecordset(sql)
  
  If Not rs.EOF Then
    Set oPx = GetPatient(rs(0), False)                  '*-- Coleman 051701
    If oPx Is Nothing Then
        LoadPatient = False
    Else
        LoadPatient = True
'''      '*-- Coleman 052301: Don't fail entire load if only sponsor fails
'''      LoadPatient = True
'''      LoadSponsor oPx, oPx.SponsorSSN                   '*-- Coleman 051701
'''      '*-- Coleman 060601: Test for empty Sponsor...should NEVER be nothing.
'''      'If oPx.Sponsor Is Nothing Then
'''      If oPx.Sponsor.SSN = "" Then
'''      '*-- Coleman 052301: Don't fail entire load if only sponsor fails
'''        Debug.Print "LoadSponsor returned empty object in LoadPatient."
'''      '  LoadPatient = False
'''      'Else
'''      '  LoadPatient = True
'''      End If
    End If
  End If
  If rs.State > 0 Then rs.Close
  Set rs = Nothing
  Set oPx = Nothing
  Debug.Print "End LoadPatient: Unit_Number=" & NCID
  
End Function

Private Function GetDeersID(ByVal DeersID As String) As String
    Dim id As String

    If LenB(DeersID) > 1 Then
        If Left$(DeersID, 1) = "D" Then
            id = Mid(DeersID, 2)
        End If
    End If

    GetDeersID = id
End Function


Public Function GetPatient(NCID As Long, Optional Additional As Boolean) As GEMS_MMI

'- Additional: is this a primary request, where I set the curr patient, or is
'- this just an 'additional' request?
'-- Additional information:  All patients have a sponsor.  If the patitient is
'-- the eligible member, then the patient's sponsor is himself (blah!)
'*-- Additional=True means GetPatient will instantiate a new GEMS_MMI object,
'*-- will load and return it with sponsor info. It does not change
'*-- the calling GEMS_MMI object.

  Dim RecCnt As Long
  Dim oPx As GEMS_MMI
  'Static GettingSponsor As Boolean     '*-- Coleman 051601: Remove recursion
  Dim sql As String
  Dim rs As Adodb.Recordset
  Dim sSSN As String
  Dim sIDTypeNcid As String
  Dim sID As String
  
  'cmd.CommandText = "Select * from Patient where ncid  = " & NCID
  
  Debug.Print "GetPatient: NCID=" & NCID
  
  sql = "SELECT * FROM mmi_id" _
    & " WHERE mmi_id.unit_number =" & NCID
  Set rs = mobjDAS.OpenRecordset(sql)
  If rs Is Nothing Then
    Exit Function
  End If
  If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Exit Function
  End If
  
  Debug.Print "   GetPatient   Unit_Number:" & rs("Unit_Number") & " - SSN:" & rs("SSN")
  
  If rs!UpdateFlag = "D" Then Exit Function
  
  Set oPx = New GEMS_MMI
  
  If Not Additional Then
    Set mobjCurrPat = oPx
    'Set oPx.Sponsor = New GEMS_MMI     '*-- Coleman 051601: Remove recursion
  End If
  
  oPx.NCID = rs("unit_number")
  oPx.UnitNumber = rs("unit_number")
  
'2/21/02 - Walker - Does not catch Null date values
'  If Not IsNull(rs("birth_date")) Then
'  If rs("birth_date") <> "12/31/9999" Then
    If IsNull(rs!birth_date) Then
        oPx.BirthDate = cwdNULL_DATE
    Else
        oPx.BirthDate = "" & rs("Birth_Date")
    End If
'  End If
  oPx.MaritalStatus = "" & rs("Marital_Status_cid")
 ' oPx.PCM = "" & rs("PCM")
 ' oPx.PCMAddress = "" & rs("PCMAddress")
 ' oPx.PCMLocation = "" & rs("PCMLocation")
 ' oPx.PCMNCID = "" & rs("PCMNCID")
 ' oPx.PCMSpecialty = "" & rs("PCMSpecialty")
  oPx.Race = "" & rs("Race_cid")
  oPx.Religion = "" & rs("Religion_cid")
  oPx.SEX = "" & rs("sex")
  oPx.SSN = "" & rs("ssn")
  oPx.ForeignID_Flag = "" & IIf(IsNull(rs("foreign_id_flag")), False, rs("foreign_id_flag"))
  oPx.DeployLoc = "" & rs!DeployLocation
  oPx.DeployUnit = "" & rs!DeployUnit
  oPx.DNBI = "" & rs!DNBI
'  oPx.MilitaryStatus = "" & rs("MilitaryStatus")

  '*-- Coleman 051701: Added support for patient flags
  oPx.AdoptionFlag = "" & rs("Adoption_Flag")
  oPx.ExpiredFlag = "" & rs("Expired_Flag")
  oPx.GuarantorFlag = "" & rs("Guarantor_Flag")
  oPx.OrganDonorFlag = "" & rs("Organ_Donor_Flag")
  oPx.PatientFlag = "" & rs("Patient_Flag")
  oPx.UnidentifiedFlag = "" & rs("Unidentified_Flag")
  oPx.DeersID = GetDeersID(rs("REFERENCE"))
  
  rs.Close
  
'================ NAME
  sql = "SELECT * FROM mmi_name" _
    & " WHERE unit_number =" & NCID
  Set rs = mobjDAS.OpenRecordset(sql)
  
  Dim oName As GEMS_Name
  Do While Not rs.EOF
    If Not "" & rs!UpdateFlag = "D" Then
      Set oName = New GEMS_Name
      oName.Degree = "" & rs("Degree")
      oName.Last_name = "" & rs("last_name")
      oName.First_name = "" & rs("first_name")
      oName.Middle_name = "" & rs("middle_name")
      oName.Suffix = "" & rs("Suffix")
      oName.Title = "" & rs("Title")
      oName.Name_type_cid = Val("" & rs("name_type_cid"))
'      oName.UpdateFlag = "" & rs!UpdateFlag
      '*-- Coleman 071201: Save this name if it is the preferred name or if no
      '*-- name has been saved yet at all.
      If (oPx.PatientName Is Nothing) Or _
                (oName.Name_type_cid = NCID_CURRENT_NAME) Then
        Set oPx.PatientName = oName
      End If
      oPx.Names.AddName oName
    End If
    
    rs.MoveNext
  Loop
  rs.Close
  
'================
  
  Set rs = mobjDAS.OpenRecordset("SELECT * FROM mmi_address" _
    & " WHERE mmi_address.unit_number =" & NCID)
    
Dim oAddr As GEMS_Address
  Do While Not rs.EOF
    If Not "" & rs!UpdateFlag = "D" Then
      Set oAddr = New GEMS_Address
      oAddr.AddrType = Val("" & rs("address_type_cid"))
      oAddr.Street1 = "" & rs("street_1")
      oAddr.Street2 = "" & rs("street_2")
      oAddr.City = "" & rs("city")
      oAddr.County = "" & rs("county")
      oAddr.State = "" & rs("state")
      oAddr.Zip = "" & rs("post_code")
      oAddr.Country = "" & rs("country")
'      oAddr.UpdateFlag = "" & rs!UpdateFlag
      oPx.Addresses.AddAddress oAddr
      Set oAddr = Nothing
    End If
    rs.MoveNext
  Loop
    
  rs.Close
'================ Phone
  
  Set rs = mobjDAS.OpenRecordset("SELECT * FROM mmi_phone" _
    & " WHERE unit_number = " & NCID)
    
Dim oPhone As GEMS_Phone
  Do While Not rs.EOF
    If Not "" & rs!UpdateFlag = "D" Then
      Set oPhone = New GEMS_Phone
      oPhone.PhoneType = Val("" & rs("phone_type_cid"))
      oPhone.CountryCode = "" & rs("country_code")
      oPhone.AreaCode = "" & rs("area_code")
      oPhone.Phone = "" & rs("local_number")
      oPhone.Extension = "" & rs("internal_extension")
'      oPhone.UpdateFlag = "" & rs!UpdateFlag
      oPx.phones.AddPhone oPhone
      Set oPhone = Nothing
    End If
    rs.MoveNext
  Loop
  rs.Close
   
'================ Generic id
  
  Set rs = mobjDAS.OpenRecordset("SELECT * From mmi_generic_id" _
    & " WHERE unit_number = " & NCID)

  Dim oGenID As GEMS_Generic_Id
  Do While Not rs.EOF
      '*-- Skip over records marked for deletion.
      If Not "" & rs!UpdateFlag = "D" Then
        sIDTypeNcid = Val("" & rs("id_type_ncid"))
        sID = Trim$("" & rs("id_value"))
        Set oGenID = New GEMS_Generic_Id
        oGenID.assigning_authority = Val("" & rs("assigning_authority_ncid"))
        oGenID.id_type_ncid = sIDTypeNcid
        oGenID.id_value = sID
'           oGenID.UpdateFlag = "" & rs!UpdateFlag
        oGenID.UpdateFlag = "U"     '*-- Reading data from LDDB so do
                                    '*-- update when you save...
        oPx.Generic_IDs.AddGenericID oGenID
      End If
      rs.MoveNext
  Loop
  Set oGenID = Nothing
  rs.Close
  
'-- get med rec num
  Set rs = mobjDAS.OpenRecordset("Select * from mmi_site" & _
                                    " where unit_number = " & NCID)
  If Not rs.EOF Then
    oPx.MedRecordNumber = rs!MEDICAL_RECORD_NUMBER & ""
    '*-- Coleman 100201
    oPx.FacilityNCID = "" & rs!FACILITY_CID
    oPx.Sites.Add oPx.UpdateFlag, rs!FACILITY_CID & "", _
                    rs!MEDICAL_RECORD_NUMBER & "", Val(rs!RADIOLOGY_NUMBER & "")
  End If
  rs.Close
  
  '*-- Get Generic Data
  Set rs = mobjDAS.OpenRecordset("Select * from mmi_genericdata" _
  & " where unit_number = " & NCID)

  
Dim oGenD As GEMS_Generic_Data
  Do While Not rs.EOF
    If Not "" & rs!UpdateFlag = "D" Then
      Set oGenD = New GEMS_Generic_Data
      oGenD.Ordinal = Val("" & rs("ordinal"))
      oGenD.Error_Text = "" & rs("error_text")
      oGenD.Parent = Val("" & rs("parent"))
      oGenD.Type_NCID = Val("" & rs("type_ncid"))
      oGenD.Value_NCID = Val("" & rs("Value_NCID"))
      oGenD.Value_RSID = Val("" & rs("Value_RSID"))
      oGenD.Value_Text = "" & rs("Value_Text")
      oGenD.UpdateFlag = "U"
      Select Case rs!Type_NCID
        Case NCID_COMMAND_INTEREST
            '*-- Coleman 111901: Filter out duplicate entries
            'oPx.Command_Interest.Add rs!Value_Text
            On Error Resume Next
            oPx.Command_Interest.Add rs!Value_Text, rs!Value_Text
            On Error GoTo 0
        Case NCID_COMMAND_SECURITY
            '*-- Coleman 111901: Filter out duplicate entries
            'oPx.Command_Security.Add rs!Value_Text
            On Error Resume Next
            oPx.Command_Security.Add rs!Value_Text, rs!Value_Text
            On Error GoTo 0
        Case NCID_SWS_FLYING_STATUS, NCID_SWS_JUMPING_STATUS, _
                NCID_SWS_DIVING_STATUS, NCID_SWS_ON_MOBILITY, _
                NCID_SWS_MP, NCID_SWS_SUBMARINE
            On Error Resume Next
            oPx.Special_Work_Status.Add rs!Value_Text, CStr(rs!Type_NCID)
            On Error GoTo 0
      End Select
      oPx.Generic_Datas.AddGenericData oGenD
    End If
    If Val("" & rs!Ordinal) > oPx.NextGDOrdinal Then oPx.NextGDOrdinal = Val("" & rs!Ordinal)
    rs.MoveNext
  Loop
  
    '*-- Coleman 051601: Remove recursion - next IF
    'If Not GettingSponsor And (Not oPx.Sponsor Is Nothing) Then
    '    If oPx.Sponsor.SSN = "" Then
    '        Set oPx.Sponsor = oPx
    '        oPx.Sponsor.SSN = oPx.SSN
    '    End If
    'End If
   
   '*-- Now that you are done getting data objects, inc NextGDOrdinal
   '*-- to guarantee it is at least one greater that the last value used.
   oPx.NextGDOrdinal = oPx.NextGDOrdinal + 1
   
 Set GetPatient = oPx
  Set oPx = Nothing
  If rs.State > 0 Then rs.Close
  Set rs = Nothing
End Function

Private Function PatientRecExists(Optional lUnitNum As Variant) As Boolean
    Dim rs As Adodb.Recordset
    Dim sSQL As String
    Dim lUnitNumber As Long
    
    On Error GoTo UnexpectedError
    If IsMissing(lUnitNum) Then
        lUnitNumber = mobjCurrPat.UnitNumber
    Else
        lUnitNumber = lUnitNum
    End If
    sSQL = "select unit_number from mmi_id where unit_number = " & lUnitNumber
    Set rs = mobjDAS.OpenRecordset(sSQL)
    PatientRecExists = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
UnexpectedError:
    mobjShared.ShowVBError 0, "Unexpected error: Recordset open failed with Unit_Number=" & CStr(rs!Unit_Number), "GEMS_MMIOps.PatientRecExists", "Patient Error", vbExclamation
    Exit Function
End Function

'    T.Noren  This Procedure rewritten on 04/02/01
'             to remove all references to record sets
'             converting updates and new record additions
'             to SQL statments
'
'   Coleman 092901: Formatting changes only.
'
'-- save the patient to the database
'
' WRogers 10/03/2003
' FOREIGN_ID_FLAG added to the database schema to describe whether
' the SSN field contains a US or non US id. Sponsor is assumed the
' attribute is the whatever the SSN is.
'
Public Function PutPatient(Optional oPx As GEMS_MMI) As Boolean

    Dim sql As String
    Dim sSetString As String
'    Dim NCID As Long
    Dim WhereClause As String
    Dim Values As String
    Dim ActionClause As String
    Dim oSeqId As ICHCSII_SQL.ISqlOps
    Dim ParseChar As String
    Dim oDatabase As CHCSII_CONN.Conn
    
    Set oDatabase = New CHCSII_CONN.Conn
    'Dim oRepToNCID As GEMS_RepToNCID
    
    If oPx Is Nothing Then
        Set oPx = mobjCurrPat
    End If
    
    If oPx Is Nothing Then Exit Function
    
    On Error GoTo ErrHandler
    
    'WR added to support various data backends
    If mobjShared.UseMySQLDB Then
        geDBUsed = edb_MySQL
    ElseIf mobjShared.UseLocalDB Then
        If mobjShared.UseSQLServer Then
            geDBUsed = edb_sqlserver
        Else
            geDBUsed = edb_Oracle
        End If
    Else
        geDBUsed = edb_Oracle
    End If
    
    
    ParseChar = Chr$(255)
    ' Begin:SCR #35666;   Developer: Mag 05/07/2003 04:36 PM
    'Removed NCID declaration since oPx.NCID will work in its place.
    'When oPx.NCID = "" then we received an error Type mismatch
'    NCID = oPx.NCID
    ' End: SCR #35666;
    mobjDAS.BeginTrans
     
    '*-- Coleman 052501: DataManager blindly saves patient info and will pass
    '*-- in the new flag "?" to have me check if the record should be updated
    '*-- or inserted.
    If oPx.UpdateFlag = "?" Or oPx.UpdateFlag = "" Then
        If PatientRecExists(Val(oPx.NCID)) Then
            oPx.UpdateFlag = "U"
        Else
            oPx.UpdateFlag = "N"
        End If
    End If
  
    If oPx.UpdateFlag = "D" Then
        '*--
        '*-- Delete from all MMI tables for this unit_number and exit.
        '*--
        sSetString = "set updateflag = 'D' where unit_number = " & NCID
        mobjDAS.ExecuteSQL "Update mmi_id " & sSetString
        mobjDAS.ExecuteSQL "Update mmi_name " & sSetString
        mobjDAS.ExecuteSQL "Update mmi_address " & sSetString
        mobjDAS.ExecuteSQL "Update mmi_phone " & sSetString
        mobjDAS.ExecuteSQL "Update mmi_genericdata " & sSetString
        mobjDAS.ExecuteSQL "Update mmi_generic_id " & sSetString
        '*-- Coleman 052401: Added support of site
        mobjDAS.ExecuteSQL "Update mmi_site " & sSetString
        mobjDAS.CommitTrans
        
        PutPatient = True
        Exit Function
      
    '*--
    '*-- Update or Insert in MMI_ID
    '*--
    ElseIf oPx.UpdateFlag = "N" Or (oPx.UpdateFlag = "" And oPx.NCID = "") Then
        oPx.UpdateFlag = "N"
        ActionClause = "INSERT INTO"
    ElseIf oPx.UpdateFlag = "U" Or (oPx.UpdateFlag = "" And oPx.NCID <> "") Then
        oPx.UpdateFlag = "U"
        ActionClause = "Update"
        WhereClause = "WHERE Unit_Number = " & oPx.NCID
    End If
  
    If ActionClause <> "" Then
        If oPx.NCID = "" Then
            Set oSeqId = oDatabase.CHCSII_SQLOPS 'New GEMS_SQL.SqlOps
            oPx.NCID = oSeqId.GetNextNcid
        End If
        
        If oPx.BirthDate > 0 Then
            'SCR 23520 MEDCIN tool, Encounter Training tool
            Values = Values & "Birth_Date = " & SQLDate(oPx.BirthDate) & ParseChar
        End If
        If Trim$(oPx.MaritalStatus) <> "" Then
            Values = Values & "Marital_Status_CID = " & _
                        RTrim$(oPx.MaritalStatus) & "" & ParseChar
        End If
        If Trim$(oPx.Race) <> "" Then
            Values = Values & "Race_cid = " & Trim$(oPx.Race) & "" & ParseChar
        End If
        If Trim$(oPx.Religion) <> "" Then
            Values = Values & "Religion_CID = " & Trim$(oPx.Religion) & _
                        "" & ParseChar
        End If
        If Trim$(oPx.SEX) <> "" Then
            Values = Values & "Sex = '" & Trim$(oPx.SEX) & "'" & ParseChar
        End If
'        If Trim(oPx.ForeignID) <> "" Then
'            Values = Values & "ForeignID = '" & Trim$(oPx.ForeignID) & "'" & ParseChar
'            Values = Values & "SSN = '" & Trim$(oPx.SSN) & "'" & ParseChar
'        Else
'            If Trim$(oPx.SSN) <> "" Then
'                Values = Values & "ForeignID = '" & Trim$(oPx.ForeignID) & "'" & ParseChar
'                Values = Values & "SSN = '" & Trim$(oPx.SSN) & "'" & ParseChar
'            End If
'        End If
        If Trim(oPx.SSN) <> "" Then
            Values = Values & "Foreign_ID_Flag = " & CInt(oPx.ForeignID_Flag) & ParseChar
            Values = Values & "SSN = '" & Trim$(oPx.SSN) & "'" & ParseChar
        End If
        If Trim$(oPx.DeployUnit) <> "" Then
            '*-- Coleman 111401: SCR 18746
            'Values = Values & "DeployUnit = '" & Trim$(oPx.DeployUnit) & "'" & ParseChar
            Values = Values & "DeployUnit = " & SQLQuote(Trim$(oPx.DeployUnit)) & ParseChar
        End If
        If Trim$(oPx.DeployLoc) <> "" Then
            'Values = Values & "DeployLocation = '" & Trim$(oPx.DeployLoc) & "'" & ParseChar
            Values = Values & "DeployLocation = " & SQLQuote(Trim$(oPx.DeployLoc)) & ParseChar
        End If
        If Trim$(oPx.DNBI) <> "" Then
            'Values = Values & "DNBI = '" & Trim$(oPx.DNBI) & "'" & ParseChar
            Values = Values & "DNBI = " & SQLQuote(Trim$(oPx.DNBI)) & ParseChar
        End If
        If Trim$(oPx.AdoptionFlag) <> "" Then
            Values = Values & "Adoption_Flag = '" & Trim$(oPx.AdoptionFlag) & "'" & ParseChar
        End If
        If Trim$(oPx.ExpiredFlag) <> "" Then
            Values = Values & "Expired_Flag = '" & Trim$(oPx.ExpiredFlag) & "'" & ParseChar
        End If
        If Trim$(oPx.GuarantorFlag) <> "" Then
            Values = Values & "Guarantor_Flag = '" & Trim$(oPx.GuarantorFlag) & "'" & ParseChar
        End If
        If Trim$(oPx.OrganDonorFlag) <> "" Then
            Values = Values & "Organ_Donor_Flag = '" & Trim$(oPx.OrganDonorFlag) & "'" & ParseChar
        End If
        If Trim$(oPx.PatientFlag) <> "" Then
            Values = Values & "Patient_Flag = '" & Trim$(oPx.PatientFlag) & "'" & ParseChar
        Else
            ' Begin:SCR #39619;   Developer: Mag 08/21/2003 02:10 PM
            'force patient flag
            Values = Values & "Patient_Flag = 'Y'" & ParseChar
            ' End: SCR #39619;
        End If
        If Trim$(oPx.UnidentifiedFlag) <> "" Then
            Values = Values & "Unidentified_Flag = '" & Trim$(oPx.UnidentifiedFlag) & "'" & ParseChar
        End If
        '*-- Coleman 062101: Add data items
        If Trim$(oPx.MothersName) <> "" Then
            'Values = Values & "Mother_Name = '" & Trim$(oPx.MaidenName) & "'" & ParseChar
            Values = Values & "Mother_Name = " & SQLQuote(Trim$(oPx.MothersName)) & ParseChar
        End If
        
        'WRogers MothersMaidenName was not a synonym for MothersName
        If Trim$(oPx.MaidenName) <> "" Then
            'Values = Values & "Mother_Name = '" & Trim$(oPx.MaidenName) & "'" & ParseChar
            Values = Values & "Maiden_Name = " & SQLQuote(Trim$(oPx.MaidenName)) & ParseChar
        End If
        
        Values = Values & "Unit_Number = " & oPx.NCID & ParseChar
        '*-- Coleman 042401: Must be non-null but what really belongs here?
        If Len(oPx.DeersID) > 0 Then
            Values = Values & "Reference = 'D" & oPx.DeersID & "'" & ParseChar
        Else
            Values = Values & "Reference = '" & oPx.NCID & "'" & ParseChar
        End If
        Values = Values & "Create_Audit_Num = 576" & ParseChar
        
        '*-- No ParseChar at end of last values term.
        Values = Values & "UpdateFlag = '" & oPx.UpdateFlag & "' "
        DoWrite mobjDAS, ActionClause, WhereClause, "MMI_ID", Values, ParseChar
     
    End If
  
    '*--
    '*-- Update or Insert to MMI_NAME
    '*--
    Dim oName As GEMS_Name
    Dim iIndex As Double
    
    iIndex = 1
    
    For Each oName In oPx.Names
        
        ActionClause = ""
        WhereClause = ""
        Values = ""
        
'    If oPx.PatientName.UpdateFlag = "D" Then
'        sql = "Update MMI_Name " & sSetString & " And Name_type_cid = " & oPx.PatientName.Name_type_cid
'        mobjDAS.ExecuteSQL sql
'    Else
    
        'avoid constrant violations, check the address type before insert
        'WR 4/4/2002
        If mobjDAS.OpenRecordset("Select count(*) num from MMI_Name Where Unit_Number = " & oPx.NCID _
                                    & " And Name_Type_cid = " & oName.Name_type_cid).Fields("num") > 0 Then
                    
            oName.UpdateFlag = "U"
        Else
            oName.UpdateFlag = "N"
        End If
    
    
        If oName.UpdateFlag = "N" Or oPx.UpdateFlag = "N" Then
            ActionClause = "INSERT INTO"
            oName.UpdateFlag = "N"
        ElseIf oPx.PatientName.UpdateFlag = "U" Then
           ActionClause = "UPDATE"
           WhereClause = "WHERE unit_number = " & oPx.NCID & " and Name_type_cid = " & CStr(oName.Name_type_cid)
        End If
        
        '*-- Coleman 111401: SCR 18601 (added SQLQuote call 6 places below)
        If ActionClause <> "" Then
            With oName
                If Trim$(.Degree) <> "" Then
                    Values = Values & "Degree = " & SQLQuote(Trim$(.Degree)) & ParseChar
                End If
                If Trim$(.Last_name) <> "" Then
                    Values = Values & "last_name = " & SQLQuote(Trim$(.Last_name)) & ParseChar
                End If
                If Trim$(.First_name) <> "" Then
                    Values = Values & "first_name = " & SQLQuote(Trim$(.First_name)) & ParseChar
                End If
                If Trim$(.Middle_name) <> "" Then
                    Values = Values & "middle_name = " & SQLQuote(Trim$(.Middle_name)) & ParseChar
                End If
                If Trim$(.Suffix) <> "" Then
                    Values = Values & "Suffix = " & SQLQuote(Trim$(.Suffix)) & ParseChar
                End If
                If Trim$(.Title) <> "" Then
                    Values = Values & "Title = " & SQLQuote(Trim$(.Title)) & ParseChar
                End If
                '*-- Coleman 042401: Make sure Name_Type_CID gets something.
                '*-- If .Name_type_cid <> "" Then Values = Values & "name_type_cid = " & .Name_type_cid & "" & ParseChar
                If .Name_type_cid <> "" Then
                    Values = Values & "name_type_cid = " & .Name_type_cid & "" & ParseChar
                Else
                    Values = Values & "name_type_cid = " & NCID_CURRENT_NAME & "" & ParseChar
                End If
                
                '*-- Coleman 062101:
                Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                Values = Values & "Unit_Number = " & oPx.NCID & ParseChar
                Values = Values & "Name_Index = " & iIndex & "" & ParseChar
                Values = Values & "updateflag = '" & .UpdateFlag & "' "
                
                DoWrite mobjDAS, ActionClause, WhereClause, "MMI_Name", Values, ParseChar
            End With
            iIndex = iIndex + 1
        End If
        
    Next oName
'    End If
  
    '*--
    '*-- Update or Insert to MMI_ADDRESS
    '*--
    Dim oAddr As GEMS_Address
    
    For Each oAddr In oPx.Addresses
    
        ActionClause = ""
        WhereClause = ""
        Values = ""
        
        With oAddr
        
            If .UpdateFlag = "D" Then
                sql = "delete from MMI_Address " & sSetString & " where Address_Type_cid = " & .AddrType
                mobjDAS.ExecuteSQL sql
            Else
                'avoid constrant violations, check the address type before insert
                'WR 4/4/2002
                If mobjDAS.OpenRecordset("Select count(*) num from MMI_Address Where Unit_Number = " & oPx.NCID _
                                            & " And Address_Type_cid = " & .AddrType).Fields("num") > 0 Then
                            
                    .UpdateFlag = "U"
                Else
                    .UpdateFlag = "N"
                End If
                
                If .UpdateFlag = "N" Then
                    ActionClause = "INSERT INTO"
                ElseIf .UpdateFlag = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "Where Unit_Number = " & oPx.NCID & " and Address_Type_cid = " & .AddrType
                End If
                
                '*-- Coleman 111401: SCR 18746 -- Added calls to SQLQuote
                If ActionClause <> "" Then
                    If Trim$(.AddrType) <> "" Then
                        Values = Values & "address_type_cid = " & Trim$(.AddrType) & "" & ParseChar
                    End If
                    If Trim$(.Street1) <> "" Then
                        'Values = Values & "street_1 = '" & Trim$(.Street1) & "'" & ParseChar
                        Values = Values & "street_1 = " & SQLQuote(Trim$(.Street1)) & ParseChar
                    End If
                    If Trim$(.Street2) <> "" Then
                        'Values = Values & "street_2 = '" & Trim$(.Street2) & "'" & ParseChar
                        Values = Values & "street_2 = " & SQLQuote(Trim$(.Street2)) & ParseChar
                    End If
                    If Trim$(.City) <> "" Then
                        'Values = Values & "city = '" & Trim$(.City) & "'" & ParseChar
                        Values = Values & "city = " & SQLQuote(Trim$(.City)) & ParseChar
                    End If
                    If Trim$(.County) <> "" Then
                        'Values = Values & "county = '" & Trim$(.County) & "'" & ParseChar
                        Values = Values & "county = " & SQLQuote(Trim$(.County)) & ParseChar
                    End If
                    If Trim$(.State) <> "" Then
                        'Values = Values & "state = '" & Trim$(.State) & "'" & ParseChar
                        Values = Values & "state = " & SQLQuote(Trim$(.State)) & ParseChar
                    End If
                    If Trim$(.Zip) <> "" Then
                        Values = Values & "post_code = '" & Trim$(.Zip) & "'" & ParseChar
                    End If
                    If Trim$(.Country) <> "" Then
                        'Values = Values & "country = '" & Trim$(.Country) & "'" & ParseChar
                        Values = Values & "country = " & SQLQuote(Trim$(.Country)) & ParseChar
                    End If
                    '*-- Coleman 062101:
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    Values = Values & "Unit_Number = " & oPx.NCID & ParseChar
                    Values = Values & "updateflag = '" & .UpdateFlag & "' "
                    
                    DoWrite mobjDAS, ActionClause, WhereClause, "MMI_Address", Values, ParseChar
                End If
            End If
        End With
    Next oAddr
        
    '*--
    '*-- Update or Insert to MMI_PHONE
    '*--
    Dim oPhone As GEMS_Phone
 
    For Each oPhone In oPx.phones
        ActionClause = ""
        WhereClause = ""
        Values = ""
  
        With oPhone
            If .UpdateFlag = "D" Then
                sql = "delete from MMI_Phone " & sSetString & " where phone_type_cid = " & .PhoneType
                mobjDAS.ExecuteSQL sql
            Else
                'WR
                '4/4/2002
                If mobjDAS.OpenRecordset("Select count(*) num from MMI_Phone Where Unit_Number = " & oPx.NCID _
                                            & " And Phone_Type_cid = " & .PhoneType).Fields("num") > 0 Then
                            
                    .UpdateFlag = "U"
                Else
                    .UpdateFlag = "N"
                End If

            
                If .UpdateFlag = "N" Then
                    ActionClause = "INSERT INTO"
                ElseIf .UpdateFlag = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "WHERE unit_number = " & oPx.NCID & " and phone_type_cid = " & .PhoneType
                End If
                
                ' SCR #31073;   Developer:  11/18/2003
                ' Country_code, Area_code, Internal_extension statements need single quotes
                If ActionClause <> "" Then
                    If Trim$(.PhoneType) <> "" Then
                        Values = Values & "phone_type_cid = " & Trim$(.PhoneType) & "" & ParseChar
                    End If
                    If Trim$(.CountryCode) <> "" Then
                        Values = Values & "country_code = '" & Trim$(.CountryCode) & "' " & ParseChar
                    End If
                    If Trim$(.AreaCode) <> "" Then
                        Values = Values & "area_code = '" & Trim$(.AreaCode) & "' " & ParseChar
                    End If
                    If Trim$(.Phone) <> "" Then
                        Values = Values & "local_number = '" & Trim$(.Phone) & "' " & ParseChar
                    End If
                    If Trim$(.Extension) <> "" Then
                        Values = Values & "internal_extension = '" & Trim$(.Extension) & "' " & ParseChar
                    End If
                    '*-- Coleman 062101:
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    Values = Values & "Unit_Number = " & oPx.NCID & ParseChar
                    Values = Values & "updateflag = '" & .UpdateFlag & "' "
                    
                    DoWrite mobjDAS, ActionClause, WhereClause, "MMI_Phone", Values, ParseChar
                End If
            End If
        End With
      
    Next oPhone
    
    '*--
    '*-- Update or Insert to MMI_GENERIC_ID
    '*--
    Dim oGenI As GEMS_Generic_Id
    
    For Each oGenI In oPx.Generic_IDs
        ActionClause = ""
        WhereClause = ""
        Values = ""
        
        With oGenI
            If .UpdateFlag = "D" Then
                sql = "delete from mmi_generic_id " & sSetString _
                    & " where id_Type_ncid = " & CStr(.id_type_ncid)
                mobjDAS.ExecuteSQL sql
            Else
                If mobjDAS.OpenRecordset("Select count(*) num from mmi_generic_id Where Unit_Number = " & oPx.NCID _
                                            & " and id_Type_ncid = " & CStr(.id_type_ncid)).Fields("num") > 0 Then
                            
                    .UpdateFlag = "U"
                Else
                    .UpdateFlag = "N"
                End If

            
                If .UpdateFlag = "N" Then
                    ActionClause = "INSERT INTO"
                ElseIf .UpdateFlag = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "WHERE Unit_Number = " & oPx.NCID & " and id_Type_ncid = " & CStr(.id_type_ncid)
                End If
                
                If ActionClause <> "" Then
                    If .assigning_authority > 0 Then
                        Values = Values & "assigning_authority_ncid = " & CStr(.assigning_authority) & "" & ParseChar
                    End If
                    If .id_type_ncid > 0 Then
                        Values = Values & "id_type_ncid = " & CStr(.id_type_ncid) & "" & ParseChar
                    End If
                    If Trim$(.id_value) <> "" Then
                        'Values = Values & "id_value = '" & .id_value & "'" & ParseChar
                        Values = Values & "id_value = " & SQLQuote(.id_value) & ParseChar
                    End If
                    '*-- Coleman 062101: Bug fix - null audit field breaks appointments.
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    Values = Values & "Unit_Number = " & oPx.NCID & ParseChar
                    Values = Values & "updateflag = '" & .UpdateFlag & "' "
                    
                    DoWrite mobjDAS, ActionClause, WhereClause, "MMI_Generic_ID", Values, ParseChar
                End If
            End If
        End With
    Next oGenI
     
    '*--
    '*-- Update or Insert to MMI_GENERICDATA
    '*--
    Dim oGenD As GEMS_Generic_Data
    
    For Each oGenD In oPx.Generic_Datas
        ActionClause = ""
        WhereClause = ""
        Values = ""
        
        With oGenD
            Dim rs As Adodb.Recordset
            
            If .UpdateFlag = "D" Then
                sql = "delete from MMI_GenericData " & sSetString _
                    & " where ordinal= " & .Ordinal
                mobjDAS.ExecuteSQL sql
            Else
                Set rs = mobjDAS.OpenRecordset("Select count(*) num, ordinal, Value_NCID from mmi_genericdata Where Unit_Number = " & oPx.NCID _
                                & " and Type_ncid = " & CStr(.Type_NCID) & " group by ordinal, Value_NCID")
                                
                'mmi_genericdata contains possibly more than one command interest and
                '   command security
                'WR 5/1/2002 scr 18847
                If rs.RecordCount Then
                    If .Type_NCID = NCID_COMMAND_INTEREST Or _
                        .Type_NCID = NCID_COMMAND_SECURITY Then
                        
                        Do While Not rs.EOF
                            If .Value_NCID = rs("Value_NCID").Value Then
                                .Ordinal = rs("ordinal").Value
                                Exit Do
                            End If
                            rs.MoveNext
                        Loop
                    Else
                        If rs.Fields("num") > 0 Then
                            .Ordinal = rs("ordinal").Value
                            .UpdateFlag = "U"
                        End If
                    End If
                End If
            
                If .UpdateFlag = "N" Or rs.RecordCount = 0 Then
                    'need to set a default value for ordinal
                    'If oSeqId Is Nothing Then Set oSeqId = New GEMS_SQL.SqlOps
                    
                    Set rs = mobjDAS.OpenRecordset("Select max(ordinal) maximum from mmi_genericdata Where Unit_Number = " & oPx.NCID _
                                & " ")

                    If rs.RecordCount = 0 Then
                        .Ordinal = 1
                    
                    'WR problem occurs on new patients downloaded
                    ElseIf IsNull(rs("maximum")) Then
                        .Ordinal = 1
                    Else
                        .Ordinal = rs("maximum").Value + 1
                    End If
                    
                    ActionClause = "INSERT INTO"
                ElseIf .UpdateFlag = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "WHERE unit_number = " & oPx.NCID & " and ordinal= " & .Ordinal
                End If
                
                If ActionClause <> "" Then
                    If .Ordinal > 0 Then
                        Values = Values & "ordinal = " & CStr(.Ordinal) & "" & ParseChar
                    End If
                    If Trim$(.Error_Text) <> "" Then
                        'Values = Values & "error_text = '" & Trim$(.Error_Text) & "'" & ParseChar
                        Values = Values & "error_text = " & SQLQuote(Trim$(.Error_Text)) & ParseChar
                    End If
                    If .Parent > 0 Then
                        Values = Values & "parent = " & CStr(.Parent) & "" & ParseChar
                    End If
                    If .Type_NCID > 0 Then
                        Values = Values & "type_ncid = " & CStr(.Type_NCID) & "" & ParseChar
                    End If
                    '*-- Coleman 112601: SCR 17043 - Facility NCID can be < 0...
                    If .Value_NCID <> 0 Then
                        Values = Values & "Value_NCID = " & CStr(.Value_NCID) & "" & ParseChar
                    End If
                'SCRs 19698, 19700 and 19701 - 2/28/02 - Walker - Cannot clear data with this code
                '2/28/02    If Trim$(.Value_Text) <> "" Then
                        Values = Values & "Value_text = " & SQLQuote(Trim$(.Value_Text)) & ParseChar
                '2/28/02    End If
                    If .Value_RSID > 0 Then
                        Values = Values & "Value_RSID = " & CStr(.Value_RSID) & "" & ParseChar
                    End If
                    '*-- Coleman 062101:
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    Values = Values & "Unit_Number = " & oPx.NCID & ParseChar
                    Values = Values & "updateflag = '" & .UpdateFlag & "' "
                    
                    DoWrite mobjDAS, ActionClause, WhereClause, "MMI_GenericData", Values, ParseChar
                End If
            End If
        End With
      
    Next oGenD
    
    '*--
    '*-- Update or Insert to MMI_SITE
    '*--
    '*-- Coleman 042501: Added support of Site Information
    Dim oSite As GEMS_Site
    Dim lCnt As Long
    
    '*-- Coleman 101901: Transfer of patients to the LDDB from
    '*-- CDR requires that we put in a SITE record, so make one
    '*-- up if none exist...Needed to create new appointments.
    If oPx.Sites.Count = 0 And oPx.UpdateFlag = "N" Then
        oPx.SaveSiteObj "", oPx.MedRecordNumber
    End If
    
    'SCR 25089 WR
    For lCnt = 1 To oPx.Sites.Count
    'For lCnt = 1 To oPx.Sites.Count
    'For Each oSite In oPx.Sites
        Set oSite = oPx.Sites.Item(lCnt)
        
        ActionClause = ""
        WhereClause = ""
        Values = ""
        
        With oSite
            '*-- No UpdateFlag in this table, so do nothing:
            If .UpdateFlag = "D" Then
                sql = "delete from MMI_Site " & sSetString & " where Facility_cid=" & CLng(.FacilityNCID)
                mobjDAS.ExecuteSQL sql
            Else
                If mobjDAS.OpenRecordset("Select count(*) num  from mmi_site Where Unit_Number = " & oPx.NCID _
                                            & " and facility_cid = " & CStr(.FacilityNCID)).Fields("num") > 0 Then
                            
                    .UpdateFlag = "U"
                Else
                    .UpdateFlag = "N"
                End If
            
                If .UpdateFlag = "N" Then
                    ActionClause = "INSERT INTO"
                ElseIf .UpdateFlag = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "WHERE unit_number = " & oPx.NCID & _
                                    " AND facility_cid = " & CStr(.FacilityNCID)
                End If
                
                If ActionClause <> "" Then
                    '*-- To Do: Patient.Facility is a text string, we need real CID for laptop...
                    If Trim$(.FacilityNCID) <> "" Then
                        Values = Values & "FACILITY_CID = " & _
                                    CStr(.FacilityNCID) & ParseChar
                    End If
                    
                    If Trim$(.MedRecNum) <> "" Then
                        'Values = Values & "MEDICAL_RECORD_NUMBER = '" _
                                & Trim$(.MedRecNum) & "' " & ParseChar
                        Values = Values & "MEDICAL_RECORD_NUMBER = " _
                                & SQLQuote(Trim$(.MedRecNum)) & ParseChar
                    End If
                    
                    If Trim$(.RadiologyNum) <> "" Then
                        'Values = Values & "RADIOLOGY_NUMBER = '" _
                                   & Trim$(.RadiologyNum) & "' " & ParseChar
                        Values = Values & "RADIOLOGY_NUMBER = " _
                                   & SQLQuote(Trim$(.RadiologyNum)) & ParseChar
                    End If
                    
                    '*-- Create_Audit_Num is arbitrary non-null value
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    '*-- Last values term added should not have ParseChar at end.
                    Values = Values & "UNIT_NUMBER = " & CStr(oPx.UnitNumber) & ""
                    
                    DoWrite mobjDAS, ActionClause, WhereClause, "MMI_Site", Values, ParseChar
                End If
            End If
        End With
    'Next oSite
    Next lCnt
    '*-- Coleman 042501: Added Site Information -- end
    
    mobjDAS.CommitTrans
    oPx.UpdateFlag = ""
    Set oSeqId = Nothing
    PutPatient = True
    Exit Function
    
ErrHandler:

    mobjDAS.RollbackTrans
    '*-- Coleman 061401: Use standard error dialog
    mobjShared.ShowVBError Err.Number, Err.Description, "GEMS_MMIOps.PutPatient", "Patient Error", vbExclamation
    Set oSeqId = Nothing
    
    Exit Function
    Resume                'Debug
    
End Function

Private Sub Class_Initialize()
  Dim oDatabase As CHCSII_CONN.Conn
  
  Set oDatabase = New CHCSII_CONN.Conn
  
  Set mobjShared = New CWShared
  
  Set mobjDAS = oDatabase.CHCSII_DAS(Auto)
  
  Set mobjCurrPat = New GEMS_MMI
  Set mobjCurrPat.Sponsor = New GEMS_MMI
  
End Sub

Public Sub LoadSponsor(oPx As GEMS_MMI, SponsorSSN As String)
    '*-- Pass in the patient who wants to have his sponsor loaded.
    '-- get the sponsor as described by the ssn
    Dim rs As Adodb.Recordset
    Dim sSQL As String
    
    Debug.Print "LoadSponsor: " & SponsorSSN
    '*-- Coleman 060701: Update SponsorSSN with new value you want to load.
    oPx.SponsorSSN = SponsorSSN
    If SponsorSSN = "" Then
        '*-- Always load a sponsor object, even if it is empty.
        Set oPx.Sponsor = New GEMS_MMI
    Else
        sSQL = "SELECT unit_number from mmi_id WHERE ssn = '" & SponsorSSN & "'"
        Set rs = mobjDAS.OpenRecordset(sSQL)
        If Not rs.EOF Then
            '*-- The SET statement actually puts the new GEMS_MMI into oPx.
            Set oPx.Sponsor = GetPatient(rs(0), True)
        Else
            '*-- Coleman 060601: Add empty Sponsor object when sponsor not in DB.
            Set oPx.Sponsor = New GEMS_MMI
        End If
        oPx.SponsorSSN = SponsorSSN
        rs.Close
        Set rs = Nothing
    End If
End Sub


'================================ current patient object ========================
'*-- Coleman 051801: Added new patient flag properties.
Public Property Get AdoptionFlag() As String
      If Not mobjCurrPat Is Nothing Then AdoptionFlag = mobjCurrPat.AdoptionFlag
End Property
Public Property Let AdoptionFlag(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.AdoptionFlag = RHV
End Property

Public Property Get ExpireFlag() As String
      If Not mobjCurrPat Is Nothing Then ExpireFlag = mobjCurrPat.ExpiredFlag
End Property
Public Property Let ExpireFlag(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.ExpiredFlag = RHV
End Property

Public Property Get GuarantorFlag() As String
      If Not mobjCurrPat Is Nothing Then GuarantorFlag = mobjCurrPat.GuarantorFlag
End Property
Public Property Let GuarantorFlag(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.GuarantorFlag = RHV
End Property

Public Property Get OrganDonorFlag() As String
      If Not mobjCurrPat Is Nothing Then OrganDonorFlag = mobjCurrPat.OrganDonorFlag
End Property
Public Property Let OrganDonorFlag(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.OrganDonorFlag = RHV
End Property

Public Property Get PatientFlag() As String
      If Not mobjCurrPat Is Nothing Then PatientFlag = mobjCurrPat.PatientFlag
End Property
Public Property Let PatientFlag(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.PatientFlag = RHV
End Property

Public Property Get UnidentifiedFlag() As String
      If Not mobjCurrPat Is Nothing Then UnidentifiedFlag = mobjCurrPat.UnidentifiedFlag
End Property
Public Property Let UnidentifiedFlag(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.UnidentifiedFlag = RHV
End Property

'*-- Coleman 062101: Adding two new properties for PIC data
Public Property Get AssigningAuthority() As String
      If Not mobjCurrPat Is Nothing Then AssigningAuthority = mobjCurrPat.AssigningAuthority
End Property
Public Property Let AssigningAuthority(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.AssigningAuthority = RHV
End Property
Public Property Get MothersName() As String
      If Not mobjCurrPat Is Nothing Then MothersName = mobjCurrPat.MothersName
End Property
Public Property Let MothersName(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.MothersName = RHV
End Property
Public Property Get MaidenName() As String
      If Not mobjCurrPat Is Nothing Then MaidenName = mobjCurrPat.MaidenName
End Property
Public Property Let MaidenName(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.MaidenName = RHV
End Property

Public Property Get DeployLoc() As String
      If Not mobjCurrPat Is Nothing Then DeployLoc = mobjCurrPat.DeployLoc
End Property
Public Property Let DeployLoc(ByVal RHV As String)
      If Not mobjCurrPat Is Nothing Then mobjCurrPat.DeployLoc = RHV
End Property

Public Property Get DeployUnit() As String
    If Not mobjCurrPat Is Nothing Then DeployUnit = mobjCurrPat.DeployUnit
End Property
Public Property Let DeployUnit(ByVal RHV As String)
    If Not mobjCurrPat Is Nothing Then mobjCurrPat.DeployUnit = RHV
End Property

Public Property Get DNBI() As String
    If Not mobjCurrPat Is Nothing Then DNBI = mobjCurrPat.DNBI
End Property
Public Property Let DNBI(ByVal RHV As String)
    If Not mobjCurrPat Is Nothing Then mobjCurrPat.DNBI = RHV
End Property

Public Property Get EOF() As Boolean
 EOF = mvarEOF
End Property
Public Property Get UnitNumber() As Long
  If Not mobjCurrPat Is Nothing Then UnitNumber = mobjCurrPat.UnitNumber
End Property
Public Property Let UnitNumber(ByVal RHV As Long)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.UnitNumber = RHV
 End Property
Public Property Get BirthDate() As Date
   If Not mobjCurrPat Is Nothing Then BirthDate = mobjCurrPat.BirthDate
End Property
Public Property Let BirthDate(ByVal RHV As Date)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.BirthDate = RHV
End Property
Public Property Get MaritalStatus() As String
  If Not mobjCurrPat Is Nothing Then MaritalStatus = mobjCurrPat.MaritalStatus
End Property
Public Property Let MaritalStatus(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.MaritalStatus = RHV
End Property
Public Property Get NCID() As String
  If Not mobjCurrPat Is Nothing Then NCID = mobjCurrPat.NCID
End Property
Public Property Let NCID(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.NCID = RHV
End Property
Public Property Get PCM() As String
   If Not mobjCurrPat Is Nothing Then PCM = mobjCurrPat.PCM
End Property
Public Property Let PCM(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.PCM = RHV
End Property
Public Property Get PCMAddress() As String
  If Not mobjCurrPat Is Nothing Then PCMAddress = mobjCurrPat.PCMAddress
End Property
Public Property Let PCMAddress(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.PCMAddress = RHV
 End Property
Public Property Get PCMDutyPhone() As GEMS_Phone
  If Not mobjCurrPat Is Nothing Then Set PCMDutyPhone = mobjCurrPat.PCMDutyPhone
End Property
Public Property Set PCMDutyPhone(ByVal RHV As GEMS_Phone)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.PCMDutyPhone = RHV
End Property
Public Property Get PCMLocation() As String
  If Not mobjCurrPat Is Nothing Then PCMLocation = mobjCurrPat.PCMLocation
End Property
Public Property Let PCMLocation(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.PCMLocation = RHV
End Property
Public Property Get PCMNCID() As String
   If Not mobjCurrPat Is Nothing Then PCMNCID = mobjCurrPat.PCMNCID
End Property
Public Property Let PCMNCID(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.PCMNCID = RHV
End Property
Public Property Get PCMSpecialty() As String
  If Not mobjCurrPat Is Nothing Then PCMSpecialty = mobjCurrPat.PCMSpecialty
End Property
Public Property Let PCMSpecialty(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.PCMSpecialty = RHV
End Property
Public Property Get Race() As String
  If Not mobjCurrPat Is Nothing Then Race = mobjCurrPat.Race
End Property
Public Property Let Race(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.Race = RHV
End Property
Public Property Get Religion() As String
  If Not mobjCurrPat Is Nothing Then Religion = mobjCurrPat.Religion
End Property
Public Property Let Religion(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.Religion = RHV
End Property
Public Property Get SEX() As String
   If Not mobjCurrPat Is Nothing Then SEX = mobjCurrPat.SEX
End Property
Public Property Let SEX(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.SEX = RHV
End Property
Public Property Get SSN() As String
   If Not mobjCurrPat Is Nothing Then SSN = mobjCurrPat.SSN
End Property
Public Property Let SSN(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.SSN = RHV
End Property
Public Property Get FacilityNCID() As String
  If Not mobjCurrPat Is Nothing Then FacilityNCID = mobjCurrPat.FacilityNCID
End Property
Public Property Let FacilityNCID(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.FacilityNCID = RHV
End Property
Public Property Get EncounterNumber() As String
  If Not mobjCurrPat Is Nothing Then EncounterNumber = mobjCurrPat.EncounterNumber
End Property
Public Property Let EncounterNumber(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.EncounterNumber = RHV
End Property
Public Property Get EncounterFacilityNCID() As String
  If Not mobjCurrPat Is Nothing Then EncounterFacilityNCID = mobjCurrPat.EncounterFacilityNCID
End Property
Public Property Let EncounterFacilityNCID(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.EncounterFacilityNCID = RHV
End Property
Public Property Get MilitaryStatus() As String
  If Not mobjCurrPat Is Nothing Then MilitaryStatus = mobjCurrPat.MilitaryStatus
End Property
Public Property Let MilitaryStatus(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.MilitaryStatus = RHV
End Property
Public Property Get MedicalRecordNumber() As String
  If Not mobjCurrPat Is Nothing Then MedicalRecordNumber = mobjCurrPat.MedRecordNumber
End Property
Public Property Let MedicalRecordNumber(ByVal RHV As String)
  If Not mobjCurrPat Is Nothing Then mobjCurrPat.MedRecordNumber = RHV
End Property
Public Property Get Addresses() As GEMS_Addresses
  If Not mobjCurrPat Is Nothing Then Set Addresses = mobjCurrPat.Addresses
End Property
Public Property Set Addresses(ByVal RHV As GEMS_Addresses)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.Addresses = RHV
End Property
Public Property Get phones() As GEMS_Phones
  If Not mobjCurrPat Is Nothing Then Set phones = mobjCurrPat.phones
End Property
Public Property Set phones(ByVal RHV As GEMS_Phones)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.phones = RHV
 End Property
Public Property Get PatientName() As GEMS_Name
   If Not mobjCurrPat Is Nothing Then Set PatientName = mobjCurrPat.PatientName
End Property
Public Property Set PatientName(ByVal RHV As GEMS_Name)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.PatientName = RHV
End Property
Public Property Get Sponsor() As GEMS_MMI
  If Not mobjCurrPat Is Nothing Then Set Sponsor = mobjCurrPat.Sponsor
End Property
Public Property Set Sponsor(ByVal RHV As GEMS_MMI)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.Sponsor = RHV
End Property
Public Property Get HomePhone() As GEMS_Phone
  If Not mobjCurrPat Is Nothing Then Set HomePhone = mobjCurrPat.HomePhone
End Property
Public Property Set HomePhone(ByVal RHV As GEMS_Phone)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.HomePhone = RHV
End Property
Public Property Get HomeAddress() As GEMS_Address
  If Not mobjCurrPat Is Nothing Then Set HomeAddress = mobjCurrPat.HomeAddress
End Property
Public Property Set HomeAddress(ByVal RHV As GEMS_Address)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.HomeAddress = RHV
End Property

Public Property Get WorkPhone() As GEMS_Phone
  If Not mobjCurrPat Is Nothing Then Set WorkPhone = mobjCurrPat.WorkPhone
End Property
Public Property Set WorkPhone(ByVal RHV As GEMS_Phone)
  If Not mobjCurrPat Is Nothing Then Set mobjCurrPat.WorkPhone = RHV
End Property

Private Sub Class_Terminate()
  Set mobjCurrPat = Nothing
  Set mobjDAS = Nothing
  Set mcolPatients = Nothing
  Set mobjCurrPat = Nothing
 
End Sub

Public Property Get UpdateFlag() As String:           UpdateFlag = mobjCurrPat.HomeAddress.UpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mobjCurrPat.HomeAddress.UpdateFlag = RHV:         End Property

Public Property Get AddrType() As String:           AddrType = mobjCurrPat.HomeAddress.AddrType:    End Property
Public Property Let AddrType(ByVal RHV As String):  mobjCurrPat.HomeAddress.AddrType = RHV:         End Property

Public Property Get Street1() As String:            Street1 = mobjCurrPat.HomeAddress.Street1:  End Property
Public Property Let Street1(ByVal RHV As String):   mobjCurrPat.HomeAddress.Street1 = RHV:      End Property

Public Property Get Street2() As String:            Street2 = mobjCurrPat.HomeAddress.Street2:  End Property
Public Property Let Street2(ByVal RHV As String):   mobjCurrPat.HomeAddress.Street2 = RHV:      End Property

Public Property Get City() As String:               City = mobjCurrPat.HomeAddress.City:        End Property
Public Property Let City(ByVal RHV As String):      mobjCurrPat.HomeAddress.City = RHV:         End Property

Public Property Get State() As String:              State = mobjCurrPat.HomeAddress.State:      End Property
Public Property Let State(ByVal RHV As String):     mobjCurrPat.HomeAddress.State = RHV:        End Property

Public Property Get Zip() As String:                Zip = mobjCurrPat.HomeAddress.Zip:          End Property
Public Property Let Zip(ByVal RHV As String):       mobjCurrPat.HomeAddress.Zip = RHV:          End Property

Public Property Get County() As String:             County = mobjCurrPat.HomeAddress.County:    End Property
Public Property Let County(ByVal RHV As String):    mobjCurrPat.HomeAddress.County = RHV:       End Property

Public Property Get Country() As String:            Country = mobjCurrPat.HomeAddress.Country:  End Property
Public Property Let Country(ByVal RHV As String):   mobjCurrPat.HomeAddress.Country = RHV:      End Property

Public Sub UpdateGenericData(sNCID As String, sValue As String, Optional sValueNCID As String)
  Call mobjCurrPat.UpdateGenericData(sNCID, sValue, sValueNCID)
End Sub

Public Sub UpdateGenericID(sNCID As String, sValue As String)
  Call mobjCurrPat.UpdateGenericID(sNCID, sValue)
End Sub

Public Property Get FMP() As String:               FMP = mobjCurrPat.FMP:            End Property
Public Property Let FMP(RHV As String):            mobjCurrPat.FMP = RHV:            End Property

Public Property Get Rank() As String:               Rank = mobjCurrPat.Rank:         End Property
Public Property Let Rank(RHV As String):            mobjCurrPat.Rank = RHV:          End Property

Public Property Get ACV_ENROLLMENT_STATUS() As String:    ACV_ENROLLMENT_STATUS = mobjCurrPat.ACV_ENROLLMENT_STATUS: End Property
Public Property Let ACV_ENROLLMENT_STATUS(RHV As String): mobjCurrPat.ACV_ENROLLMENT_STATUS = RHV:                   End Property

Public Property Get Category() As String
'SCRs 19698, 19700 and 19701 - 2/19/02 - Walker - Used In Gems
'Demographics references this as the Patient Category value
    Category = mobjCurrPat.Category
End Property
Public Property Get UIC() As String:    UIC = mobjCurrPat.UIC:        End Property
Public Property Let UIC(RHV As String): mobjCurrPat.UIC = RHV:             End Property

Public Property Get EnrollmentFacility() As String:    EnrollmentFacility = mobjCurrPat.EnrollmentFacility:        End Property
Public Property Let EnrollmentFacility(RHV As String): mobjCurrPat.EnrollmentFacility = RHV:             End Property

Public Property Get FullName() As String:    FullName = mobjCurrPat.FullName:        End Property

Public Property Get SponsorSSN() As String
'*-- Coleman 071301: Get SponsorSSN from patient's property, not the sponsor obj.
'  If Not mobjCurrPat.Sponsor Is Nothing Then
'    SponsorSSN = mobjCurrPat.Sponsor.SSN
'  End If
    SponsorSSN = mobjCurrPat.SponsorSSN
End Property
Public Property Let SponsorSSN(RHV As String)
'*-- Coleman 071301: Update SponsorSSN in patient property, not the sponsor obj.
'  If Not mobjCurrPat.Sponsor Is Nothing Then
'     mobjCurrPat.Sponsor.SSN = RHV
'  End If
    mobjCurrPat.SponsorSSN = RHV
End Property
    
'*-- Coleman 061401: Added new property
Public Property Get FMPSponsorSSN() As String
  FMPSponsorSSN = mobjCurrPat.FMPSponsorSSN
End Property
Public Property Let FMPSponsorSSN(RHV As String)
  mobjCurrPat.FMPSponsorSSN = RHV
End Property
    
'*-- Coleman 061401: Added new property
Public Property Get SponsorSrchString() As String
  SponsorSrchString = mobjCurrPat.SponsorSrchString
End Property
Public Property Let SponsorSrchString(RHV As String)
  mobjCurrPat.SponsorSrchString = RHV
End Property
    
    
'*-- Coleman 061401: Added new property
Public Property Get PatientSrchString() As String
  PatientSrchString = mobjCurrPat.PatientSrchString
End Property
Public Property Let PatientSrchString(RHV As String)
  mobjCurrPat.PatientSrchString = RHV
End Property
    
    
Public Property Get SWSDiving() As String
  SWSDiving = mobjCurrPat.SWSDiving
End Property
Public Property Let SWSDiving(ByVal RHV As String)
  mobjCurrPat.SWSDiving = RHV
End Property

Public Property Get SWSFlying() As String
  SWSFlying = mobjCurrPat.SWSFlying
End Property
Public Property Let SWSFlying(ByVal RHV As String)
  mobjCurrPat.SWSFlying = RHV
End Property

Public Property Get SWSJumping() As String
  SWSJumping = mobjCurrPat.SWSJumping
End Property
Public Property Let SWSJumping(ByVal RHV As String)
  mobjCurrPat.SWSJumping = RHV
End Property

Public Property Get SWSMP() As String
  SWSMP = mobjCurrPat.SWSMP
End Property
Public Property Let SWSMP(ByVal RHV As String)
  mobjCurrPat.SWSMP = RHV
End Property

Public Property Get SWSOnMobility() As String
  SWSOnMobility = mobjCurrPat.SWSOnMobility
End Property
Public Property Let SWSOnMobility(ByVal RHV As String)
  mobjCurrPat.SWSOnMobility = RHV
End Property

Public Property Get SWSPRP() As String
  SWSPRP = mobjCurrPat.SWSPRP
End Property
Public Property Let SWSPRP(ByVal RHV As String)
 mobjCurrPat.SWSPRP = RHV
End Property

Public Property Get SWSSubmarine() As String
  SWSSubmarine = mobjCurrPat.SWSSubmarine
End Property
Public Property Let SWSSubmarine(ByVal RHV As String)
  mobjCurrPat.SWSSubmarine = RHV
End Property


Public Property Get DEERSEligibility() As String
  DEERSEligibility = mobjCurrPat.DEERSEligibility
End Property
Public Property Let DEERSEligibility(ByVal RHV As String)
   mobjCurrPat.DEERSEligibility = RHV
End Property

Public Property Get DEERSEligibilityEnd() As String
  DEERSEligibilityEnd = mobjCurrPat.DEERSEligibilityEnd
End Property
Public Property Let DEERSEligibilityEnd(ByVal RHV As String)
  mobjCurrPat.DEERSEligibilityEnd = RHV
End Property

Public Property Get DEERSEligibilityStart() As String
  DEERSEligibilityStart = mobjCurrPat.DEERSEligibilityStart
End Property
Public Property Let DEERSEligibilityStart(ByVal RHV As String)
   mobjCurrPat.DEERSEligibilityStart = RHV
End Property

'*-- Coleman 110501: SCR 18301
Public Property Get EMailAddress() As String
  EMailAddress = mobjCurrPat.EMailAddress
End Property
Public Property Let EMailAddress(ByVal RHV As String)
   mobjCurrPat.EMailAddress = RHV
End Property

Public Property Get Records_Maintained_At() As String
  Records_Maintained_At = mobjCurrPat.Records_Maintained_At
End Property
Public Property Let Records_Maintained_At(ByVal RHV As String)
  mobjCurrPat.Records_Maintained_At = RHV
End Property

Public Property Get Comments() As String
  Comments = mobjCurrPat.Comments
End Property
Public Property Let Comments(ByVal RHV As String)
  mobjCurrPat.Comments = RHV
End Property

Public Property Get Command_Interest() As Collection
  Set Command_Interest = mobjCurrPat.Command_Interest
End Property
Public Property Let Command_Interest(ByVal RHV As Collection)
'///  Set mobjCurrPat.Command_Interest = RHV
End Property

Public Property Get Command_Security() As Collection
  Set Command_Security = mobjCurrPat.Command_Security
End Property
Public Property Let Command_Security(ByVal RHV As Collection)
'//  Set mobjCurrPat.Command_Security = RHV
End Property

'*--
'*-- Coleman 051701: Not sure if we need this and what exactly it is used for.
'*--
'Public Function LoadSponsorSSN(SSN As String) As Boolean
''- called by the child patient to load it's parent (sponsor)
'  Dim rs As ADODB.Recordset
'  Dim lSponsorUnitNum As Long           '*-- Coleman 051701
'
'  Debug.Print "GEMS_MMIOps::LoadSponsorSSN: SSN=" & SSN
'
'  If Trim(SSN) = "" Then Exit Function
'
'  Set rs = mobjDAS.OpenRecordset("Select unit_number from mmi_id where ssn = '" & SSN & "'")
'
'  If Not rs.EOF Then
'  '- set the current patient in this ops module = to the sponsor
'     Call GetPatient(CLng("" & rs(0)), lSponsorUnitNum, False)
'     LoadSponsorSSN = True
'  End If
'  rs.Close
'  Set rs = Nothing
'
'End Function
Public Property Get ForeignID_Flag() As Boolean

    ForeignID_Flag = mobjCurrPat.ForeignID_Flag
    
End Property
Public Property Let ForeignID_Flag(newId As Boolean)

    mobjCurrPat.ForeignID_Flag = newId
    
End Property

Public Property Get NextGDOrdinal() As Long
  If Not mobjCurrPat Is Nothing Then
    NextGDOrdinal = mobjCurrPat.NextGDOrdinal
  End If
End Property

Public Property Let NextGDOrdinal(RHV As Long)
  If Not mobjCurrPat Is Nothing Then
    mobjCurrPat.NextGDOrdinal = RHV
  End If
End Property

Public Function GetGenericData(ByVal sNCID As String) As String
    If Not mobjCurrPat Is Nothing Then
        GetGenericData = mobjCurrPat.GetGenericData(sNCID)
    End If
End Function

Public Function GetGenericID(ByVal sNCID As String) As String
    If Not mobjCurrPat Is Nothing Then
        GetGenericID = mobjCurrPat.GetGenericID(sNCID)
    End If
End Function

'*-- Coleman 090401: Support of GetProperty function.
Public Function GetGenericDataValueNCID(ByVal sNCID As String) As String
    If Not mobjCurrPat Is Nothing Then
        GetGenericDataValueNCID = mobjCurrPat.GetGenericDataValueNCID(sNCID)
    End If
End Function

'*-- Coleman 090701: SCR 16395
Public Function GetFMPSelectStrings(sNumNCID As String, sTextNCID As String) As Adodb.Recordset
    Dim rs As Adodb.Recordset
    Dim sSQL As String
    
    Set rs = New Adodb.Recordset
    sSQL = "select n.ncid FMP_Ncid, " & SQLTrim("n.description") & " FMP_Num, " & SQLTrim("t.Description") & " FMP_Text"
    sSQL = sSQL & " from ncids n, ncids t"
    sSQL = sSQL & " where n.domainncid = '" & sNumNCID & "'"
    sSQL = sSQL & " and t.domainncid = '" & sTextNCID & "'"
    sSQL = sSQL & " and n.ncid = t.ncid"
    
    Set rs = mobjDAS.OpenRecordset(sSQL)
    Set GetFMPSelectStrings = rs
    Set rs = Nothing

End Function
Public Function GetPatCatSelectStrings(sDomainNCID As String) As Adodb.Recordset
'SCRs 19698, 19700 and 19701 - 2/15/02 - Walker - Used In Gems - Begin
Dim rs      As Adodb.Recordset
Dim sSQL    As String
    
    Set rs = New Adodb.Recordset
    
    sSQL = "select n.ncid PatCat_Ncid, " & SQLTrim("n.Description") & " PatCat_Text"
    sSQL = sSQL & " from ncids n"
    sSQL = sSQL & " where n.domainncid = " & sDomainNCID & ""
    sSQL = sSQL & " order by n.description"
    
    Set rs = mobjDAS.OpenRecordset(sSQL)
    Set GetPatCatSelectStrings = rs
    Set rs = Nothing

End Function
