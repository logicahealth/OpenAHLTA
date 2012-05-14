VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SectionDataSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iSectionDataMapper

Private mSectionDataRS As ADODB.Recordset
Private mNotesDataRS As ADODB.Recordset

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private mMode As CHCSII_CONN.eBackend
Private mUseCache As Boolean

Private Const mstrDegreeTag     As String = "<%%degree tag%%>"
Private Const EMPTY_CODE As String = "0"
Private Const m_strListToolRecordColumns As String = "DataId, EncounterNumber, FacilityNCID, RTF, UpdateFlag, " & _
                                                     "CreatedBy, CreatedOn, UpdatedBy, UpdatedOn, SOURCENCID, ORIGINALSIZE "

'added SOURCENCID to SQL list WR 2/14/2002
Private Const m_strListToolRecordsColumns As String = "DataId, UpdateFlag, SnoID, Duration, Modifier, Onset, " & _
                                                      "Prefix, QLink, QLinkseq, RangeNormalHigh, RangeNormalLow, " & _
                                                      "RangeScale, ReferID, Result, Status, Unit, Value, Note, ChartFlag "

Private Sub PopulateSectionDataRS(ByVal EncounterNumber As String, ByVal FacilityNCID As String)

    Dim objDAS As ICHCSII_DAS
    Dim sSql As String

    Set objDAS = mConn.CHCSII_DAS(Auto)
    
    sSql = "Select * from Enc_Sections where encounternumber = " & EncounterNumber _
      & " and facilityNCID = " & FacilityNCID
          
    Set mSectionDataRS = objDAS.OpenRecordset(sSql)
    
    Set objDAS = Nothing

End Sub

Private Sub PopulateNotesDataRS(ByVal EncounterNumber As String, ByVal FacilityNCID As String)

    Dim objDAS As ICHCSII_DAS
    Dim sSql As String

    Set objDAS = mConn.CHCSII_DAS(Auto)
    
    sSql = "Select * from Enc_rtfs where encounternumber = " & EncounterNumber _
      & " and facilityNCID = " & FacilityNCID
      
    Set mNotesDataRS = objDAS.OpenRecordset(sSql)
    
    Set objDAS = Nothing

End Sub

Private Sub iSectionDataMapper_Init(Conn As CHCSII_CONN.Conn, Cache As DL_Support.CacheMapper, ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mMode = DataAccessMode
  mUseCache = CacheConfigItems
End Sub

Private Function iSectionDataMapper_RetrieveAP(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean

    Dim oRecord As MedcinRecord
    Dim sName As String
    Dim sValue As String
    Dim i As Integer
    Dim EncDataId As Long
    Dim LTRDataID As Long
    Dim sql As String
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Dim dDTS As Date
    Dim sOwnerNCID As String
    
    Dim AnP As AnP
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set AnP = Encounter.Sections("AnP")
    
    Set Section = AnP
    
    AnP.Status = Incomplete
    Section.DataID = 0
    
    LoadSectionDataSQL eap, Section, Encounter
    
    AnP.Status = StatusNCIDToEnum(Val("" & mSectionDataRS!Status))
            
    If Not IsNull(mSectionDataRS!UPDATED_FLAG) Then
        AnP.Updated = CBool(mSectionDataRS!UPDATED_FLAG)
    End If
    
    Set AnP = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.RetrieveAnP SQL", Err.Description

End Function

Private Function iSectionDataMapper_RetrieveAutocites(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim oConcept As New GEMS_ConceptCtrl
    Dim lsize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim Index       As Integer
    Dim ChunkSize   As Long
    Dim lResult     As Long
    Dim sTxt        As String
    Dim vText As Variant

    
    Dim Autocites As Autocites
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Autocites = Encounter.Sections("Autocites")
    
    Set Section = Autocites

    Section.Document.LastModifiedDate = Now
    Section.DataID = 0
    
    LoadSectionDataSQL eAutoCite, Section, Encounter

    Autocites.LabRTF = "" & mSectionDataRS("LabRTF")
    Autocites.RadRTF = "" & mSectionDataRS("RadRTF")

    Set Autocites = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve Dental Autocites", Err.Description

End Function

Private Function iSectionDataMapper_RetrieveDental(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim Dental As Dental
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Dental = Encounter.Sections("Dental")
    
    Set Section = Dental

    Section.DataID = 0
    
    LoadSectionDataSQL eDental, Section, Encounter
    
    Set Dental = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve Dental SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveDispo(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim Disposition As Disposition
    Dim Section As iSection
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sSql As String
    Dim aBytes() As Byte
    
    On Error GoTo ErrHandler
    
    Set Disposition = Encounter.Sections("Disposition")
    
    Set Section = Disposition

    Disposition.Status = Incomplete
    Section.DataID = 0
    
    LoadSectionDataSQL eDisp, Section, Encounter
    
    Disposition.Status = StatusNCIDToEnum(Val("" & mSectionDataRS!Status))
            
    If Not IsNull(mSectionDataRS!UPDATED_FLAG) Then
        Disposition.Updated = CBool(mSectionDataRS!UPDATED_FLAG)
    End If
    
    Set oSQL = mConn.CHCSII_SQLOPS(Auto)
    
    'SCR 18618 WR 11/29/2001
    sSql = "select * from Enc_Disposition " _
            & "where encounternumber = " & Encounter.id _
            & "     and facilityNCID = " & Encounter.FacilityNCID
            
    oSQL.Execute sSql
    
    With Disposition
    
        If Not oSQL.EOF Then
            .WSDisp = "" & oSQL("WSDisp")
            .EandMCalc = "" & oSQL("EANDMCALC")
            .FollowupTimeFrame = "" & oSQL("FOLLOWUPTIMEFRAME")
            .FollowupComments = "" & oSQL("FOLLOWUUPCOMMENTS")
            .DiscussedComments = "" & oSQL("DISCUSSEDCOMMENT")
            .ItemsDiscussed = "" & oSQL("ITEMSDISCUSSED")
            'msDispositionText = "" & oSQL("SUPPLDISPDATA")
            .EandMNCID = "" & oSQL("NCID")
            '.EandMReviewed = ("" & oSQL("EandMReviewed") = "Y")
            'msEandMReviewedBy = "" & oSQL("EandMReviewedBy")
            'mdtEandMReviewDate = oSQL("EandMReviewedDate")
            .NCID = "" & oSQL("DispositionNCID")
            .Text = "" & oSQL("DispositionText")
            .AdminOptions = "" & oSQL("DispAdminOption")
            .DBNICategory = "" & oSQL("INJ_ILL_CATEGORY")
            .DBNICause = "" & oSQL("INJ_ILL_CAUSE")
            .DentalDispStatus = oSQL("DentalStatus")
            .DentalDispOptions = oSQL("DentalOptions")
            .DentalDispStatusDesc = oSQL("DentalStatusDesc")
            .DiscussedComments = oSQL("DentalDispComments")
            
            If mMode = LGS Then
                'Get Meta Data
                 .DispMetaData = gobjShared.CHCSConnection.ReadChunk("enc_disposition", "ENCOUNTERNUMBER = " & Encounter.id, "DISPMETADATA")
            Else
                aBytes = oSQL("DispMetaData")
                .DispMetaData = StrConv(aBytes, vbUnicode)
            End If
        End If
    
    End With
    
    Set oSQL = Nothing
    Set Disposition = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve Disposition SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveEducation(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
'    Dim Education As Education
'    Dim Section As iSection
'
'    On Error GoTo ErrHandler
'
'    Set Education = Encounter.Sections("Education")
'
'    Set Section = Education
'
'    Section.DataID = 0
'
'    LoadSectionDataSQL eEducation, Section
'
'    Set Education = Nothing
'    Set Section = Nothing
'
'    Exit Function
'
'ErrHandler:
'
'    Err.Raise Err.Number, "DL.Retrieve Education SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveEncNotes(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim EncNotes As EncNotes
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set EncNotes = Encounter.Sections("EncNotes")
    
    Set Section = EncNotes
    
    Dim objNote As Note
    
    Set EncNotes = RetrieveNotesDataSQL(eNotes, Section)

    
    Set EncNotes = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve EncNotes SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveHistory(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim History As History
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set History = Encounter.Sections("History")
    
    Set Section = History

    Section.DataID = 0
    
    LoadSectionDataSQL eHistory, Section, Encounter
    
    Set History = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve History SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveRFV(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim RFV As RFV
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set RFV = Encounter.Sections("RFV")
    
    Set Section = RFV

    Section.DataID = 0
    
    LoadSectionDataSQL eRFV, Section, Encounter
            
    If Not IsNull(mSectionDataRS!UPDATED_FLAG) Then
        RFV.Updated = CBool(mSectionDataRS!UPDATED_FLAG)
    End If
    
    Set RFV = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve RFV SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveSO(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim SO As SO
    Dim Section As iSection
    Dim oRecord          As MedcinRecord
    Dim objNote         As SONote
    Dim oConcept        As GEMS_ConceptCtrl
    Dim objISQLRecord   As ICHCSII_SQL.ISqlOps
    Dim objISQLNote     As ICHCSII_SQL.ISqlOps
    Dim strSQL          As String
    Dim oShared         As CWShared
    Dim SectionTxtNte   As iSection
    Dim sWhereClause    As String
    
    On Error GoTo ErrHandler
    
    Set oConcept = New GEMS_ConceptCtrl
    
    Set SO = Encounter.Sections("SO")
    
    Set Section = SO

    Set oShared = New CWShared
    
    'Set mcolNotes = New Collection
    Set objISQLRecord = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Set objISQLNote = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    

    If mMode <> LGS Then
        strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord " & _
                 "WHERE EncounterNumber = " & Encounter.id & " " & _
                 "AND facilityncid = " & Encounter.FacilityNCID & " " & _
                 "AND SOURCENCID = " & SectionEnumToNCID(eSO)
    Else 'SCR-48052
        strSQL = "SELECT " & m_strListToolRecordColumns & ", TEXT_NOTE_RTF FROM ListToolRecord " & _
                 "WHERE EncounterNumber = " & Encounter.id & " " & _
                 "AND facilityncid = " & Encounter.FacilityNCID & " " & _
                 "AND SOURCENCID = " & SectionEnumToNCID(eSO)
    End If
    
    
    objISQLNote.Execute strSQL
    
    Do While Not objISQLNote.EOF
    
        Set objNote = New SONote
        Set Section = objNote
        'objNote.meType = ListNoteType
        With Section
            .DataID = objISQLNote("DataID")
            
            If Not IsNull(objISQLNote("UpdatedBy")) Then
                .Document.OwnerNCID = objISQLNote("UpdatedBy")
                oConcept.uniqueID = .Document.OwnerNCID
                .Document.OwnerName = oConcept.PrefRep("2000").Representation
            ElseIf Not IsNull(objISQLNote("CreatedBy")) Then
                .Document.OwnerNCID = objISQLNote("CreatedBy")
                oConcept.uniqueID = .Document.OwnerNCID
                .Document.OwnerName = oConcept.PrefRep("2000").Representation
            End If
            
            'If Not IsNull(objISQLNote("UpdatedOn")) Then
            'WRogers 9/21/2001 SCR 16498
            If Len(objISQLNote("UpdatedOn")) Then
                .Document.LastModifiedDate = objISQLNote("UpdatedOn")
            End If
            'objNote.msTitle = "" & objRS("Title")
            
            If mMode <> LGS Then 'SCR-48052
                'Decompress if compressed SCR-24683
                If CVar(objISQLNote("OriginalSize")) > 0 Then
                    Dim objZlib As EncZLib
                    Set objZlib = New EncZLib
                    Dim aBytes() As Byte
                    
                    aBytes = objISQLNote("RTF")
                    Call objZlib.DecompressData(aBytes, objISQLNote("OriginalSize"))
                    .Document.body = StrConv(aBytes, vbUnicode)
                End If
            
            Else
                'PGUI Mode SCR#50699
                
                sWhereClause = "ENCOUNTERNUMBER = " & Encounter.id & " AND facilityncid = " & Encounter.FacilityNCID & " AND DATAID = " & Section.DataID
                
                .Document.body = ReadChunkCHCSConn("LISTTOOLRECORD", sWhereClause, "RTF", False)
                
                Set SectionTxtNte = objNote.TextNote
                
                SectionTxtNte.Document.body = ReadChunkCHCSConn("LISTTOOLRECORD", sWhereClause, "TEXT_NOTE_RTF", False)
            End If
        End With
        
         strSQL = "SELECT " & m_strListToolRecordsColumns & " FROM ListToolRecords " & _
                  "WHERE dataid = " & Section.DataID
                  
         objISQLRecord.Execute strSQL
         
         Set objNote.MedcinRecords = New Collection
         
         Do While Not objISQLRecord.EOF
             Set oRecord = New MedcinRecord
             ReadMedcinRecordFromRS oRecord, objISQLRecord
             objNote.MedcinRecords.Add oRecord
             objISQLRecord.MoveNext
         Loop
         
         SO.Add objNote, CStr(Section.DataID)
         
         objISQLNote.MoveNext
    Loop

    Set oRecord = Nothing
    Set Section = Nothing
    Set objNote = Nothing
    Set oConcept = Nothing
    Set objISQLRecord = Nothing
    Set objISQLNote = Nothing
    Set SO = Nothing

    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve SO SQL", Err.Description
End Function

Private Function iSectionDataMapper_RetrieveVitals(Encounter As DTOs.Encounter, Params As DL_Support.iParameters) As Boolean
    Dim Vitals As Vitals
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Vitals = Encounter.Sections("Vitals")
    
    Set Section = Vitals

    Section.DataID = 0
    
    LoadSectionDataSQL eVitals, Section, Encounter

    Set Vitals = Nothing
    Set Section = Nothing
    
    Exit Function
    
ErrHandler:

    Err.Raise Err.Number, "DL.Retrieve Vitals SQL", Err.Description
End Function

Private Function iSectionDataMapper_SaveAP(Encounter As DTOs.Encounter) As Boolean
' Assume section already locked for update before getting here
    Dim oRecord     As MedcinRecord
    Dim sDiagnosis As Variant
    Dim sql        As String
    Dim DataID     As String
    Dim oSQLEx       As ICHCSII_SQL.ISqlOpsEx
    Dim oSQL       As ICHCSII_SQL.ISqlOps
    Dim szSQL       As String
    Dim oSQLGen     As SQLGen
    Dim lStatus     As Long
    
    Set oSQL = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
   
    Dim AnP As AnP
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set AnP = Encounter.Sections("AnP")
    
    Set Section = AnP

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eap), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, .Document.body, "", "", StatusEnumToNCID(AnP.Status), AnP.Updated, mConn, mMode
    End With

    
    '*************************
    ' SAVE MEDCIN RECORDS
    '*************************
    '-- we don't store the dataid for list tool records. so go get it
    oSQL.Execute "select dataid from listToolRecord where encounternumber = " & Encounter.id & " and facilityncid = " & Encounter.FacilityNCID _
      & " and SourceNCID = " & SectionEnumToNCID(eap)
    If Not oSQL.EOF Then
        DataID = oSQL.Value(0)
        'DataId = oSQL("dataid")
    Else
    '-- create the parent record in listtoolrecord table
       DataID = oSQL.GetNextID()
        oSQL.Execute "Insert into listToolRecord (dataid,EncounterNumber, FacilityNCID, ClinicNCID, SourceNCID) values (" _
             & DataID & "," & Encounter.id & "," & Encounter.FacilityNCID & "," _
             & Encounter.ClinicNCID & "," & SectionEnumToNCID(eap) & ")"
    End If
    
    oSQL.Execute "delete from listToolRecords where DataID = " & DataID

    oSQL.Execute ("SELECT * FROM ListToolRecords WHERE DataID = -1")
    'oSQL.Execute "listToolRecords" 'This method doesn't work for SQL Server SCR-609 ITT
    
    For Each oRecord In AnP.Diagnoses
        
        If Val(oRecord.SnoID) > 0 Then 'SCR-40393
            szSQL = GenerateMedcinSQL(oRecord, DataID)
            
            oSQL.Execute szSQL
            
            Set oSQLGen = Nothing
        End If

    Next
    
    For Each oRecord In AnP.Procedures
        If Val(oRecord.SnoID) > 0 Then 'SCR-40393
            szSQL = GenerateMedcinSQL(oRecord, DataID)
            oSQL.Execute szSQL
            
            Set oSQLGen = Nothing
        End If
    Next

    iSectionDataMapper_SaveAP = True
    Set oSQL = Nothing
    Set oSQLEx = Nothing

    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "Encounter.SaveAP SQL DL", Err.Description
End Function

Private Function GenerateMedcinSQL(ByRef oRecord As MedcinRecord, ByVal DataID As String) As String

    Dim oSQLGen As SQLGen
    
    Set oSQLGen = New SQLGen
    
    oSQLGen.SetTransType = eSQLInsert

    With oRecord
        oSQLGen.AddToSQL "DataID", DataID, eSQLNumber
        oSQLGen.AddToSQL "updateflag", "N", eSQLString
        oSQLGen.AddToSQL "SnoID", .SnoID, eSQLString
        oSQLGen.AddToSQL "Duration", .Duration, eSQLString
        oSQLGen.AddToSQL "Modifier", .Modifier, eSQLString
        oSQLGen.AddToSQL "Onset", .Onset, eSQLString
        oSQLGen.AddToSQL "prefix", .Prefix, eSQLString
        oSQLGen.AddToSQL "qlink", .QualifierLink, eSQLString
        oSQLGen.AddToSQL "qlinkseq", .QualifierLinkSequence, eSQLNumber
        oSQLGen.AddToSQL "RangeNormalHigh", .RangeNormalHigh, eSQLNumber
        oSQLGen.AddToSQL "RangeNormalLow", .RangeNormalLow, eSQLNumber
        oSQLGen.AddToSQL "RangeScale", .RangeScale, eSQLNumber
        oSQLGen.AddToSQL "Referid", .Referid, eSQLNumber
        oSQLGen.AddToSQL "result", .result, eSQLString
        oSQLGen.AddToSQL "Status", .Status, eSQLString
        oSQLGen.AddToSQL "Unit", .Unit, eSQLString
        oSQLGen.AddToSQL "Value", .Value, eSQLString
        
        GenerateMedcinSQL = oSQLGen.GenerateSQL("ListToolRecords", "")
    End With

End Function

Private Function iSectionDataMapper_SaveAutocites(Encounter As DTOs.Encounter) As Boolean
'
'    If Not mobjEncounter.LockingObject.StartSectionUpdate(mobjParent.meSection, 0) Then
'        MsgBxARMd "Your lock on the " & SectionEnumToName(mobjParent.meSection) & " section was broken by another user, and your changes cannot be saved.", vbInformation, "Encounter"
'        GoTo Cleanup
'    End If
    
    'date:01/07/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
    
    Dim Autocites As Autocites
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Autocites = Encounter.Sections("Autocites")
    
    Set Section = Autocites

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eAutoCite), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, .Document.body, Autocites.LabRTF, Autocites.RadRTF, 0, False, mConn, mMode
    End With

    iSectionDataMapper_SaveAutocites = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Autocite SQL DL", Err.Description
Cleanup:
    'Call mobjEncounter.LockingObject.EndSectionUpdate(mobjParent.meSection, 0)
End Function

Private Function iSectionDataMapper_SaveDental(Encounter As DTOs.Encounter) As Boolean
    
    Dim Dental As Dental
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Dental = Encounter.Sections("Dental")
    
    Set Section = Dental

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eDental), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, .Document.body, vbNullString, vbNullString, 0, False, mConn, mMode
    End With

    iSectionDataMapper_SaveDental = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Dental SQL DL", Err.Description
Cleanup:
    Set Dental = Nothing
    Set Section = Nothing
End Function

Private Function iSectionDataMapper_SaveDispo(Encounter As DTOs.Encounter) As Boolean
    
    Dim Disposition     As Disposition
    Dim Section         As iSection
    Dim objSQLEx        As ICHCSII_SQL.ISqlOpsEx
    Dim objSQL          As ICHCSII_SQL.ISqlOps
    Dim sSql            As String
    Dim oSQLGen         As SQLGen
    Dim aBytes()        As Byte
    Dim szSQL           As String
    Dim sbuf            As String
    
    On Error GoTo ErrHandler
    
    Set Disposition = Encounter.Sections("Disposition")
    
    Set Section = Disposition

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eDisp), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, .Document.body, vbNullString, vbNullString, 0, False, mConn, mMode
    End With

    '*************************
    ' SAVE DISPOSITION DATA
    '*************************
    
    Set objSQLEx = mConn.CHCSII_SQLOPS(Auto)
    
    If mMode <> LGS Then
       sSql = "select * from enc_disposition " & _
                "where EncounterNumber = " & Encounter.id & _
                " and FacilityNCID = " & Encounter.FacilityNCID
                  
        Call objSQLEx.Execute(sSql)
        
        With objSQLEx
            'WRogers 9/17/2001
            'SCR 16204
            If .EOF Then
                .AddNew
                If Len(Disposition.EandMNCID) > 0 Then 'SCR-22881
                    If StrComp(Disposition.EandMNCID, EMPTY_CODE) Then 'ITT Change
                        !NCID = Disposition.EandMNCID
                    End If
                End If
            End If
            
            !EandMCalc = Disposition.EandMCalc
            !WSDisp = Disposition.WSDisp
            !EncounterNumber = Encounter.id
            If Len(Disposition.NCID) > 0 Then 'ITT Change
                !DispositionNCID = Disposition.NCID
            End If
            !DispositionText = Disposition.Text
            !DispAdminOption = Disposition.AdminOptions
            !FollowupTimeFrame = Disposition.FollowupTimeFrame
            !FollowuupComments = Disposition.FollowupComments ' check that spelling
            !DiscussedComment = Disposition.DiscussedComments
            !ItemsDiscussed = Disposition.ItemsDiscussed
            !FacilityNCID = Encounter.FacilityNCID
            !INJ_ILL_CATEGORY = Disposition.DBNICategory
            !INJ_ILL_CAUSE = Disposition.DBNICause
            'added by #28830 to support Dental requirements
            !DentalStatus = Disposition.DentalDispStatus
            !DentalOptions = Disposition.DentalDispOptions
            !DentalStatusDesc = Disposition.DentalDispStatusDesc
            !DentalDispComments = Disposition.DiscussedComments
            aBytes = StrConv(Replace(Disposition.DispMetaData, Chr(0), ""), vbFromUnicode)
            !DispMetaData = aBytes
            .Update
            'Rs.Close
        End With
        
        'Call mobjEncounter.SaveHIPAAData
        
        'May need to refresh data on Appointment screen SCR#49319
        'mobjEncounter.Comm.Message cwiREFRESH_DATA, "APPTID|" & mobjEncounter.AppointmentID, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION
    Else
    
        Set objSQL = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
        
        'mobjEncounter.SaveWAM True 'SCR#45188
        
        Set oSQLGen = New SQLGen
        
        sSql = "select * from enc_disposition " & _
                "where EncounterNumber = " & Encounter.id & _
                "   and FacilityNCID = " & Encounter.FacilityNCID
                
        Call objSQL.Execute(sSql)
        
        With objSQL
            If .EOF Then
                'Add New
                oSQLGen.SetTransType = eSQLInsert
            Else
                oSQLGen.SetTransType = eSQLUpdate
            End If
        End With
    
        With oSQLGen
            If Len(Disposition.EandMCalc) Then
                .AddToSQL "EandMCalc", Disposition.EandMCalc, eSQLString
            End If
            .AddToSQL "WSDisp", Disposition.WSDisp, eSQLString
            .AddToSQL "EncounterNumber", Encounter.id, eSQLNumber
            If Len(Disposition.NCID) Then 'TelCons do not have disp ncid SCR-40501
                .AddToSQL "DispositionNCID", Disposition.NCID, eSQLNumber
            End If
            .AddToSQL "DispositionText", Disposition.Text, eSQLString
            .AddToSQL "DispAdminOption", Disposition.AdminOptions, eSQLString
            .AddToSQL "FollowupTimeFrame", Disposition.FollowupTimeFrame, eSQLString
            .AddToSQL "FollowuupComments", Disposition.FollowupComments, eSQLString
            .AddToSQL "DiscussedComment", Disposition.DiscussedComments, eSQLString
            .AddToSQL "ItemsDiscussed", Disposition.ItemsDiscussed, eSQLString
            .AddToSQL "FacilityNCID", Encounter.FacilityNCID, eSQLNumber
            If Len(Disposition.DentalDispStatus) Then
                .AddToSQL "DentalStatus", Disposition.DentalDispStatus, eSQLNumber
            End If
            .AddToSQL "DentalOptions", Disposition.DentalDispOptions, eSQLString
            .AddToSQL "DentalStatusDesc", Disposition.DentalDispStatusDesc, eSQLString
            .AddToSQL "DentalDispComments", Disposition.DiscussedComments, eSQLString
            
            szSQL = .GenerateSQL("enc_disposition", "EncounterNumber = " & Encounter.id)
        End With
        
        objSQL.Execute szSQL
        
        'Save Meta Data
        sbuf = Disposition.DispMetaData
        gobjShared.CHCSConnection.AppendChunk "enc_disposition", "ENCOUNTERNUMBER = " & Encounter.id, "DISPMETADATA", sbuf

        Set oSQLGen = Nothing
        
    End If

    iSectionDataMapper_SaveDispo = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Dispostion SQL DL", Err.Description
Cleanup:
    Set Disposition = Nothing
    Set Section = Nothing
End Function

Private Function iSectionDataMapper_SaveEducation(Encounter As DTOs.Encounter) As Boolean
    
    Dim Education As Education
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Education = Encounter.Sections("Education")
    
    Set Section = Education

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eEducation), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, .Document.body, vbNullString, vbNullString, 0, False, mConn, mMode
    End With

    iSectionDataMapper_SaveEducation = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Education SQL DL", Err.Description
Cleanup:
    Set Education = Nothing
    Set Section = Nothing
End Function

Private Function iSectionDataMapper_SaveEncNotes(Encounter As DTOs.Encounter) As Boolean
    
    Dim bWasNew As Boolean
    Dim Note As Note
    Dim Section As iSection
    
    On Error GoTo ErrHandler
'    If objNote.mnDataID = 0 Then
'        bWasNew = True
'    Else
'        If Not mobjEncounter.LockingObject.StartSectionUpdate(eNotes, objNote.LockID) Then
'            GoTo Cleanup
'        End If
'    End If
    
    'Check if encounter owner is Dr. Unassigned SCR-24349
    'Call mobjEncounter.CheckEncounterOwnership
    
    Dim sql As String
    Dim oSQLOps As ICHCSII_SQL.ISqlOps
    Set oSQLOps = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Dim rtf_index As Long
        
    If Section.DataID = 0 Then
    
        rtf_index = oSQLOps.GetNextID()
        sql = "insert into enc_rtfs (ENCOUNTERNUMBER,FACILITYNCID,DTS, rtftype, ENC_RTFSINDEX ) "
        sql = sql & " values (" & Encounter.id & "," & Encounter.FacilityNCID & ","
        sql = sql & mConn.SQLDate(Section.Document.LastModifiedDate) & ", 'NOTE', " & rtf_index & ")"

        oSQLOps.Execute sql
        Section.DataID = rtf_index
    End If
    WriteNoteToDatabase Note
    
    iSectionDataMapper_SaveEncNotes = True

    GoTo Cleanup
ErrHandler:
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "TextNote.GEMS_Save ", "EncounterOps", vbCritical)
Cleanup:
    Set oSQLOps = Nothing
End Function

Private Function iSectionDataMapper_SaveHistory(Encounter As DTOs.Encounter) As Boolean
    
    Dim History As History
    Dim Note As Note
    Dim Section As iSection
    
    Dim sql As String
    Dim oSQLOps As ICHCSII_SQL.ISqlOps
    Set oSQLOps = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    Dim rtf_index As Long
    
    On Error GoTo ErrHandler
    
    Set History = Encounter.Sections("History")
    
    Set Section = Note

' Assume section NOT already locked for update before getting here

    Dim bWasNew As Boolean
    On Error GoTo ErrHandler
'
'    If objItem.mnDataID = 0 Then
'        bWasNew = True
'    Else
'        If Not mobjEncounter.LockingObject.StartSectionUpdate(eHistory, objItem.mnDataID) Then
'            GoTo Cleanup
'        End If
'    End If
    
    If Section.DataID = 0 Then
        bWasNew = True
        rtf_index = oSQLOps.GetNextID()
        sql = "insert into enc_rtfs (ENCOUNTERNUMBER,FACILITYNCID,DTS, rtftype, ENC_RTFSINDEX ) "
        sql = sql & " values (" & Encounter.id & "," & Encounter.FacilityNCID & ","
        sql = sql & mConn.SQLDate(Section.Document.LastModifiedDate) & ", 'TextNote', " & rtf_index & ")"
        
        oSQLOps.Execute sql
        Section.DataID = rtf_index
    End If
    
    WriteNoteToDatabase Note

    If bWasNew Then
        History.Add Note, CStr(Section.DataID)
    Else
        'Call mobjEncounter.LockingObject.EndSectionUpdate(eHistory, Section.DataId)
    End If

    iSectionDataMapper_SaveHistory = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save History SQL DL", Err.Description
Cleanup:
    Set History = Nothing
    Set Section = Nothing
End Function

Private Function iSectionDataMapper_SaveRFV(Encounter As DTOs.Encounter) As Boolean
    
    Dim Dental As Dental
    Dim Section As iSection
    
    On Error GoTo ErrHandler
    
    Set Dental = Encounter.Sections("Dental")
    
    Set Section = Dental

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eDental), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, .Document.body, vbNullString, vbNullString, 0, False, mConn, mMode
    End With

    iSectionDataMapper_SaveRFV = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Dental SQL DL", Err.Description
Cleanup:
    Set Dental = Nothing
    Set Section = Nothing
End Function

Private Function iSectionDataMapper_SaveSO(Encounter As DTOs.Encounter) As Boolean
    
    Dim SO As SO
    Dim Section As iSection
    Dim Note As SONote
    Dim objSQLEx           As ICHCSII_SQL.ISqlOpsEx
    Dim strSQL             As String
    Dim oRecord             As MedcinRecord
    
    On Error GoTo ErrHandler
    
    Set objSQLEx = mConn.CHCSII_SQLOPS_EX(Auto)
    
    Set SO = Encounter.Sections("SO")
    
    Set Section = Note

    '*************************
    ' SAVE SECTION DATA
    '*************************
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, SectionEnumToNCID(eSO), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, vbNullString, vbNullString, vbNullString, 0, False, mConn, mMode

    
    
        '*************************
        ' SAVE LISTTOOLRECORD (contains RTF)
         '*************************
        If .DataID < 1 Then
           .DataID = objSQLEx.GetNextID(exMAINSEQ) 'GEMS_Seq_NextVal("MainSEQ")
        End If
        
        strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord WHERE DataId = " & CStr(.DataID)
    End With
    
        Call objSQLEx.Execute(strSQL)
       
        With objSQLEx
            If (.BOF And .EOF) Then
               Section.DataID = objSQLEx.GetNextID(exMAINSEQ) 'GEMS_Seq_NextVal("MainSEQ")
               .AddNew
               !DataID = Section.DataID
               !UpDateFlag = "N"
               !CreatedBy = Section.Document.OwnerNCID
               !CreatedOn = Now
               !SourceNCID = SectionEnumToNCID(eSO)
            Else
               !UpDateFlag = "U"
            End If
            
            !UpdatedBy = Section.Document.OwnerNCID
            !UpdatedOn = Now
            
            !EncounterNumber = Encounter.id
            !FacilityNCID = Encounter.FacilityNCID
            'Compress RTF SCR-24683
            !OriginalSize = Len(Section.Document.body)
            Dim objZlib As EncZLib
            Dim aBytes() As Byte
            Set objZlib = New EncZLib
            
            aBytes = StrConv(Replace(Section.Document.body, Chr(0), ""), vbFromUnicode)
            objZlib.CompressData aBytes
            !Rtf = aBytes
            '!RTF = CompressRTF(Note.msRTF)
            .Update
        End With
   
    '-- Rebuild the details
    '   Mark all current as deleted
    '   add any new ones (ltr's without dataids)
    '   update any 'old' ones
    '   that leaves all the deleted ones marked as deleted
    
    objSQLEx.Execute "DELETE FROM ListToolRecords where DataID = " & Section.DataID
    
    If Not Note.MedcinRecords Is Nothing Then 'May have blank SO Note SCR#44787
        For Each oRecord In Note.MedcinRecords
    
          Call objSQLEx.Execute("SELECT " & m_strListToolRecordsColumns & " FROM ListToolRecords WHERE DataID = -1")
          
          With objSQLEx
              If (.BOF And .EOF) Then
                 .AddNew
                 !DataID = Section.DataID
                 !UpDateFlag = "N"
                 !CreatedBy = Section.Document.OwnerNCID
                 !CreatedOn = Now
              Else
                 !UpDateFlag = "Y"
              End If
                
              WriteMedcinRecordToDatabase oRecord, objSQLEx
              
        
              !UpdatedBy = Section.Document.OwnerNCID
              !UpdatedOn = Now
              
              .Update
          End With
        Next 'objLTR
    End If
    

    iSectionDataMapper_SaveSO = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Dental SQL DL", Err.Description
Cleanup:
    Set SO = Nothing
    Set Section = Nothing
End Function

Private Function iSectionDataMapper_SaveVitals(Encounter As DTOs.Encounter) As Boolean
    
    Dim Dental As Dental
    Dim Section As iSection
    Dim sVitalsText As String
    
    On Error GoTo ErrHandler
    
    Set Dental = Encounter.Sections("Dental")
    
    Set Section = Dental

    '*************************
    ' SAVE SECTION DATA
    '*************************
        
    With Section
    
        sVitalsText = Replace(.Document.body, "°", mstrDegreeTag)

        AddEncSectionData .DataID, Encounter.FacilityNCID, Encounter.id, _
                            SectionEnumToNCID(eDental), .Document.OwnerNCID, .Document.OwnerName, .Document.LastModifiedDate, sVitalsText, vbNullString, vbNullString, 0, False, mConn, mMode
    End With

    iSectionDataMapper_SaveVitals = True

    GoTo Cleanup
ErrHandler:
   Err.Raise Err.Number, "Encounter.Save Dental SQL DL", Err.Description
Cleanup:
    Set Dental = Nothing
    Set Section = Nothing
End Function


Friend Sub WriteNoteToDatabase(ByRef Note As Note)

'GEMS   objdata is a rs of enc_rtf table
    Dim sComplete As String
    Dim vChunk As Variant
    Dim objSQL As ICHCSII_SQL.ISqlOpsEx
    Dim sReadOnly As String
    Dim oSQLGen    As SQLGen
    Dim sSql       As String
    Dim objDAS           As ICHCSII_DAS
    Dim sbuf        As String
    Dim Section     As iSection

    Set Section = Note
    
    If Note.IsComplete Then
        sComplete = "Y"
    Else
        sComplete = "N"
    End If
    
    If Note.ReadOnly Then 'Save Read Only flag SCR-786 ITT
        sReadOnly = "Y"
    Else
        sReadOnly = "N"
    End If

    If mMode = GEMS Or mMode = MYSQL Or mMode = ITT Then
        Set objSQL = mConn.CHCSII_SQLOPS(Auto)
        
        With Section
            objSQL.Execute "select * from enc_rtfs where ENC_RTFSINDEX = " & .DataID
            
            If Not objSQL.EOF Then
                objSQL("UserNCID") = .Document.OwnerNCID
                objSQL("UserName") = .Document.OwnerName
                objSQL("DTS") = .Document.LastModifiedDate
                objSQL("complete") = sComplete
                objSQL("Category") = .Document.Category
                objSQL("Title") = .Document.Title
                objSQL("ReadOnly") = CStr(sReadOnly)  'SCR-22188
                'Compress RTF SCR-24683
                objSQL("OriginalSize") = Len(.Document.body)
                
                'Compress only if RTF exists SCR-573 ITT
                If Len(.Document.body) Then
                    Dim objZlib As EncZLib
                    Dim aBytes() As Byte
                    Set objZlib = New EncZLib
                    
                    aBytes = StrConv(Replace(.Document.body, Chr(0), ""), vbFromUnicode)
                    objZlib.CompressData aBytes
                    objSQL.Value("Doc") = aBytes
                End If
                'objSQL("Doc") = CompressRTF(msRTF)
              'Call objSQL.AppendAsChunk("Doc", vChunk)
                objSQL.Update
            End If
        End With
        Set objSQL = Nothing
    Else
        'PGUI MODE
        Set objDAS = mConn.CHCSII_DAS(Auto)
        Set oSQLGen = New SQLGen
        oSQLGen.SetTransType = eSQLUpdate
    
        With Section
            oSQLGen.AddToSQL "UserNCID", .Document.OwnerNCID, eSQLString
            oSQLGen.AddToSQL "UserName", .Document.OwnerName, eSQLString
            oSQLGen.AddToSQL "DTS", .Document.LastModifiedDate, eSQLDate
            oSQLGen.AddToSQL "complete", sComplete, eSQLString
            oSQLGen.AddToSQL "Category", .Document.Category, eSQLString
            oSQLGen.AddToSQL "Title", .Document.Title, eSQLString
            oSQLGen.AddToSQL "ReadOnly", CStr(sReadOnly), eSQLString
            oSQLGen.AddToSQL "Related_Section", .Document.Section, eSQLNumber  'SCR-38142
            oSQLGen.AddToSQL "OriginalSize", Len(.Document.body), eSQLNumber
            
            sSql = oSQLGen.GenerateSQL("Enc_RTFs", "ENC_RTFSINDEX = " & .DataID)
            
            objDAS.ExecuteSQL sSql
            
            sbuf = .Document.body
            
            gobjShared.CHCSConnection.AppendChunk "Enc_RTFs", "ENC_RTFSINDEX = " & .DataID, "DOC", sbuf
        End With
        Set objDAS = Nothing
    End If

End Sub

Friend Sub WriteMedcinRecordToDatabase(ByRef oRecord As MedcinRecord, ByRef objISQL As ICHCSII_SQL.ISqlOpsEx)
    
    With objISQL
      !SnoID = oRecord.SnoID
      !Duration = oRecord.Duration
      !Modifier = oRecord.Modifier
      !Onset = oRecord.Onset
      If oRecord.Prefix = "" Then oRecord.Prefix = " "
      !Prefix = oRecord.Prefix
      !qlink = oRecord.QualifierLink
      !qlinkseq = oRecord.QualifierLinkSequence
      !RangeNormalHigh = oRecord.RangeNormalHigh
      !RangeNormalLow = oRecord.RangeNormalLow
      !RangeScale = oRecord.RangeScale
      !Referid = oRecord.Referid
      !result = oRecord.result
      !Status = oRecord.Status
      !Unit = oRecord.Unit
      !Value = oRecord.Value
      !Note = oRecord.Note
      !ChartFlag = oRecord.Flag  'SCR-21653, 21648 Was not saving flag which determines whether HPI or ROS
   End With
   
End Sub

Private Function PGUI_SaveNote(ByRef Note As SONote, ByRef Encounter As Encounter) As Boolean
    Dim oSQLGen         As SQLGen
    Dim objSQL          As ICHCSII_SQL.ISqlOps
    Dim oRecord         As MedcinRecord
    Dim strSQL          As String
    Dim Section         As iSection
    Dim SectionTxtNte   As iSection
    
    On Error GoTo ErrHandler
    
    Set Section = Note
    
    Set oSQLGen = New SQLGen
    'ADD SECTION DATA
    
    With Section
        AddEncSectionData .DataID, Encounter.FacilityNCID, CStr(Encounter.id), SectionEnumToNCID(eSO), .Document.OwnerNCID, _
                            .Document.OwnerName, .Document.LastModifiedDate, vbNullString, vbNullString, vbNullString, 0, False, mConn, mMode
                            
        Set objSQL = mConn.CHCSII_SQLOPS(Auto)
        
        If .DataID = 0 Then
            .DataID = objSQL.GetNextID(exMAINSEQ)
        End If
        
        'ADD LISTTOOLRECORD DATA
        strSQL = "SELECT " & m_strListToolRecordColumns & " FROM ListToolRecord WHERE DataId = " & CStr(.DataID)
                
        Call objSQL.Execute(strSQL)
   End With
   
   With objSQL
        If (.BOF And .EOF) Then
            oSQLGen.SetTransType = eSQLInsert
            oSQLGen.AddToSQL "DataId", Section.DataID, eSQLNumber
            oSQLGen.AddToSQL "UpdateFlag", "N", eSQLString
            oSQLGen.AddToSQL "CreatedBy", Section.Document.OwnerNCID, eSQLString
            oSQLGen.AddToSQL "CreatedOn", Now, eSQLDate
            oSQLGen.AddToSQL "SourceNCID", SectionEnumToNCID(eSO), eSQLNumber
            oSQLGen.AddToSQL "FacilityNCID", Encounter.FacilityNCID, eSQLNumber
        Else
            oSQLGen.SetTransType = eSQLUpdate
            oSQLGen.AddToSQL "UpdateFlag", "Y", eSQLString
        End If

        oSQLGen.AddToSQL "UpdatedBy", Section.Document.OwnerNCID, eSQLNumber
        oSQLGen.AddToSQL "UpdatedOn", Now, eSQLDate
        oSQLGen.AddToSQL "EncounterNumber", Encounter.id, eSQLNumber
        oSQLGen.AddToSQL "OriginalSize", Len(Section.Document.body), eSQLNumber
        
        .Execute oSQLGen.GenerateSQL("LISTTOOLRECORD", "DATAID = " & Section.DataID)
        
        oSQLGen.Clear
        'ADD RTF DATA
        gobjShared.CHCSConnection.AppendChunk "LISTTOOLRECORD", "DATAID = " & Section.DataID, "RTF", Section.Document.body
        
        'ADD TEXT NOTE DATA'SCR-48052
        Set SectionTxtNte = Note.TextNote
        
        gobjShared.CHCSConnection.AppendChunk "LISTTOOLRECORD", "DATAID = " & Section.DataID, "TEXT_NOTE_RTF", SectionTxtNte.Document.body
   End With
           
   'ADD LISTTOOLRECORDS
    '-- Rebuild the details
    '   Mark all current as deleted
    '   add any new ones (ltr's without dataids)
    '   update any 'old' ones
    '   that leaves all the deleted ones marked as deleted
    
    objSQL.Execute "DELETE FROM ListToolRecords where DataID = " & Section.DataID
    
    If Not Note.MedcinRecords Is Nothing Then 'May have blank SO Note SCR#44787
        For Each oRecord In Note.MedcinRecords

            oSQLGen.SetTransType = eSQLInsert
            oSQLGen.AddToSQL "DataId", Section.DataID, eSQLNumber
            oSQLGen.AddToSQL "UpdateFlag", "N", eSQLString
            oSQLGen.AddToSQL "CreatedBy", Section.Document.OwnerNCID, eSQLNumber
            oSQLGen.AddToSQL "CreatedOn", Now, eSQLDate
            oSQLGen.AddToSQL "UpdatedBy", Section.Document.OwnerNCID, eSQLNumber
            oSQLGen.AddToSQL "UpdatedOn", Now, eSQLDate
            
            With oRecord
                oSQLGen.AddToSQL "SnoID", .SnoID, eSQLString
                oSQLGen.AddToSQL "Duration", .Duration, eSQLString
                oSQLGen.AddToSQL "Modifier", .Modifier, eSQLString
                oSQLGen.AddToSQL "Onset", .Onset, eSQLString
                If Len(.Prefix) = 0 Then .Prefix = " "
                oSQLGen.AddToSQL "prefix", .Prefix, eSQLString
                oSQLGen.AddToSQL "qlink", .QualifierLink, eSQLString
                oSQLGen.AddToSQL "qlinkseq", .QualifierLinkSequence, eSQLNumber
                oSQLGen.AddToSQL "RangeNormalHigh", .RangeNormalHigh, eSQLNumber
                oSQLGen.AddToSQL "RangeNormalLow", .RangeNormalLow, eSQLNumber
                oSQLGen.AddToSQL "RangeScale", .RangeScale, eSQLNumber
                oSQLGen.AddToSQL "Referid", .Referid, eSQLNumber
                oSQLGen.AddToSQL "result", .result, eSQLString
                oSQLGen.AddToSQL "Status", .Status, eSQLString
                oSQLGen.AddToSQL "Unit", .Unit, eSQLString
                oSQLGen.AddToSQL "value", .Value, eSQLString
                oSQLGen.AddToSQL "ChartFlag", .Flag, eSQLNumber
             End With
            objSQL.Execute oSQLGen.GenerateSQL("LISTTOOLRECORDS", "")
            oSQLGen.Clear
            
            'ADD NOTE DATA
            gobjShared.CHCSConnection.AppendChunk "LISTTOOLRECORDS", "DATAID = " & Section.DataID, "Note", oRecord.Note
    
        Next 'objLTR
    End If
    
    Set oSQLGen = Nothing
    Set objSQL = Nothing
    Set oRecord = Nothing
    
    Exit Function
    
ErrHandler:
    'Call mobjEncounter.CWShared.ShowVBError(Err.Number, Err.Description, "SO.PGUI_SaveNote", "EncounterOps", vbCritical)
End Function

Private Function LoadSectionDataSQL(ByVal SectionEnum As EncSectionEnum, ByRef Section As iSection, ByRef Encounter As Encounter) As Boolean

    Dim dDTS As Date
    Dim sOwnerNCID As String
    Dim sql As String
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
      
      
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    If mSectionDataRS Is Nothing Then
        PopulateSectionDataRS Encounter.id, Encounter.FacilityNCID
    End If

    mSectionDataRS.Filter = "ENC_SECTIONSINDEX = " & SectionEnum
    
    If Not mSectionDataRS.EOF Then mSectionDataRS.MoveFirst

    With Section
    
        .DataID = 0
        
        
        Set oSQL = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
        
        If mSectionDataRS.EOF Then
            '- make one
            .DataID = oSQL.GetNextID()
            sql = "insert into enc_sections " & _
                  "(dataid, encounternumber, facilityncid, enc_sectionsindex) values (" & _
                  " " & .DataID & ", " & _
                  " " & Encounter.id & ", " & _
                  " " & Encounter.FacilityNCID & ", " & _
                  " " & SectionEnumToNCID(SectionEnum) & ") "
                  
            oSQL.Execute sql
            
        Else
            
            
            If Not IsNull(mSectionDataRS!DTS) Then
                dDTS = CDate(mSectionDataRS!DTS)
            End If
            
            sOwnerNCID = "" & mSectionDataRS!OwnerNCID
            
            .Document.OwnerName = "" & mSectionDataRS!OwnerName
            .DataID = Val("" & mSectionDataRS!DataID)
            
            If mMode = ITT Then
                
                If Len(sOwnerNCID) And sOwnerNCID <> "0" Then
                    'Check if data has been updated SCR#43993
                    If Not (dDTS = .Document.LastModifiedDate And StrComp(sOwnerNCID, .Document.OwnerNCID, vbTextCompare) = 0) Then
                        .Document.body = ReadChunkCHCSConn("ENC_SECTIONS", "ENCOUNTERNUMBER = " & mSectionDataRS("ENCOUNTERNUMBER") & " AND enc_sectionsindex = " & SectionEnumToNCID(eap), "DOC", False)
                    End If
                End If
            Else
                'Decompress if compressed SCR-24683
                If CVar(mSectionDataRS!OriginalSize) > 0 Then
                    Set objZlib = New EncZLib
                    aBytes = mSectionDataRS!Doc
                    Call objZlib.DecompressData(aBytes, mSectionDataRS!OriginalSize)
                    .Document.body = StrConv(aBytes, vbUnicode)
                    'msRTF = DecompressRTF("" & mSectionDataRS!Doc, mSectionDataRS!OriginalSize)
                End If
            End If
            
            .Document.OwnerNCID = sOwnerNCID
            .Document.LastModifiedDate = dDTS
            
        End If
    End With

End Function

Private Function RetrieveNotesDataSQL(ByVal SectionEnum As EncSectionEnum, ByRef Section As iSection) As EncNotes

    Dim dDTS As Date
    Dim sOwnerNCID As String
    Dim EncNotes As EncNotes
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sNoteType As String
    Dim objNote As Note
    
    Set EncNotes = New EncNotes
    
    If mNotesDataRS Is Nothing Then
        'PopulateNotesDataRS
    End If
    
    Select Case SectionEnum
        Case eNotes
            sNoteType = "'NOTE'"
        Case eHistory
            sNoteType = "'TextNote'"
        Case eLabs
            sNoteType = "'LABNOTE'"
        Case eRads
            sNoteType = "'RADNOTE'"
        Case eQuestionnaire
            sNoteType = "'QNRNOTE'"
            
    End Select

    mNotesDataRS.Filter = "RTFTYPE = " & sNoteType
    
    Do While Not mNotesDataRS.EOF
        Set objNote = New Note
        ReadNoteFromRecordSet objNote, mNotesDataRS
        
        Set Section = objNote
        
        ' the events are given to use in descending order, so reverse them
        If EncNotes.Count > 0 Then
            EncNotes.Add objNote, CStr(Section.DataID), 1
        Else
            EncNotes.Add objNote, CStr(Section.DataID)
        End If
        Set objNote = Nothing
        mNotesDataRS.MoveNext
    Loop
    
    Set RetrieveNotesDataSQL = EncNotes
    
End Function

Friend Sub ReadNoteFromRecordSet(ByRef Note As Note, ByRef objRS As ADODB.Recordset)
    Dim oConcept As New GEMS_ConceptCtrl
    Dim Section As iSection
    Dim objSQL As ICHCSII_SQL.ISqlOpsEx
    Dim vChunk As Variant

'    On Error GoTo ErrHandler
    
    Set Section = Note
    
    If Not objRS.EOF Then
        With Section
            .DataID = Val("" & objRS("ENC_RTFSINDEX"))
            .Document.OwnerNCID = "" & objRS("UserNCID")
            oConcept.uniqueID = .Document.OwnerNCID
            .Document.OwnerName = oConcept.PrefRep("2000").Representation
            .Document.LastModifiedDate = "" & objRS("DTS")
            
            'msRTF = "" & objRS("Doc")
            If "" & objRS("complete") = "Y" Then
                Note.IsComplete = True
            End If
            .Document.Category = "" & objRS("Category")
            .Document.Title = "" & objRS("TITLE")
            
            If "" & objRS("ReadOnly") = "Y" Then 'Load Read Only flag SCR-786 ITT
                Note.ReadOnly = True
            End If
            
            Note.RelatedSection = Val("" & objRS("RELATED_SECTION"))
            
            If mMode = LGS Then
                'msRTF = oShared.CHCSConnection.ReadChunk("enc_rtfs", "ENC_RTFSINDEX = " & mnDataID, "DOC")
                .Document.body = ReadChunkCHCSConn("enc_rtfs", "ENC_RTFSINDEX = " & .DataID, "DOC", False) 'SCR#43993
                Note.ReadOnly = True 'SCR-39840
            Else
                'Decompress if compressed SCR-24683
                If CVar(objRS("OriginalSize")) > 0 Then
                    Dim objZlib As EncZLib
                    Set objZlib = New EncZLib
                    Dim aBytes() As Byte
                    
                    aBytes = objRS("Doc")
                    Call objZlib.DecompressData(aBytes, objRS("OriginalSize"))
                    .Document.body = StrConv(aBytes, vbUnicode)
                    'msRTF = DecompressRTF("" & objRS("DOC"), objRS("OriginalSize"))
                Else
                    Set objSQL = mConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
                    objSQL.Execute "select doc from enc_rtfs where ENC_RTFSINDEX = " & Section.DataID
                    Call objSQL.GetAsChunk("Doc", vChunk)
                    .Document.body = vChunk
                End If
                Set objSQL = Nothing
            
            End If
        End With
    End If
End Sub
