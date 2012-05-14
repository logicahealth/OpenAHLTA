VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHCSII_FDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum DrugIsActive
    Active = 1
    InActive = 0
    NotInFDBFormulary = -1
End Enum

Public Enum AllergyType_Enum
    AllergyType_AllergenGroup = 1
    AllergyType_DispensableDrug = 4
    AllergyType_Ingredient = 6
End Enum

Public Enum DrugCheckSeverityLevel_Enum
    DrugCheckSeverityLevel_Contraindicated = 1
    DrugCheckSeverityLevel_Severe = 2
    DrugCheckSeverityLevel_Moderate = 3
End Enum

Private mMedID As String
Private mMedDesc As String
Private mNDC As String
Private mDefaultRoute As String
Private mDefaultForm As String
Private mDefaultDose As String
Private mDosageStrength As String
Private mContentUnit As String

Public Property Let NDC(Value As String)
  mNDC = Value
End Property
Public Property Get NDC() As String
  NDC = mNDC
End Property
Public Property Let DefaultRoute(Value As String)
  mDefaultRoute = Value
End Property
Public Property Get DefaultRoute() As String
  DefaultRoute = mDefaultRoute
End Property

Public Property Let DefaultForm(Value As String)
  mDefaultForm = Value
End Property
Public Property Get DefaultForm() As String
  DefaultForm = mDefaultForm
End Property

Public Property Let DefaultDose(Value As String)
  mDefaultDose = Value
End Property
Public Property Get DefaultDose() As String
  DefaultDose = mDefaultDose
End Property

Public Property Let DosageStrength(Value As String)
  mDosageStrength = Value
End Property
Public Property Get DosageStrength() As String
  DosageStrength = mDosageStrength
End Property

Public Property Let ContentUnit(Value As String)
  mContentUnit = Value
End Property
Public Property Get ContentUnit() As String
  ContentUnit = mContentUnit
End Property

Public Property Let MedID(Value As String)
  mMedID = Value
End Property
Public Property Get MedID() As String
  MedID = mMedID
End Property
Public Property Let MedDesc(Value As String)
  mMedDesc = Value
End Property
Public Property Get MedDesc() As String
  MedDesc = mMedDesc
End Property

'*** Get info from Caller
Public Property Set Comm(oCWComm As Object)
  Set cwobjCWComm = oCWComm
End Property
Public Property Let NCID_CALLER_APP(Value As String)
  NCID_CALLER_APPLICATION = Value '*** GLOBAL
End Property

Public Sub SetCallInfo(oCWComm As Object, CALLER_APP_NCID As String)
    
    Set cwobjCWComm = oCWComm
    NCID_CALLER_APPLICATION = CALLER_APP_NCID

    Call SetCreateDAS

End Sub

Private Sub Class_Initialize()
    
    goFDBCount = goFDBCount + 1
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    DoEvents
    
    SetFDBConnection
    
    SetFDBNav
    
    Call SetCreateDAS
    

End Sub


Private Sub Class_Terminate()
        
    On Error Resume Next
    
    goFDBCount = goFDBCount - 1
    
    If goFDBCount = 0 Then
        gbCheckedFDB = False
        Set goFDBConn = Nothing
        Set goDAS = Nothing
        Set goFDBNav = Nothing
        Set gobjShared = Nothing
    End If
    
End Sub

Public Function Search_CHCSII_Meds(ByRef oMeds As Collection, sSearchString As String, Optional IncludeInactive As Boolean = True) As Integer
    
    Dim objRS  As Recordset
    Dim oMed As Drug2Drug.MedData

    Dim sSql As String
    
    On Error GoTo Err_Handler
    
    '*** Clear out previous Searched Result
    Set oMeds = Nothing
    Set oMeds = New Collection
    
    Call SetCreateDAS
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    If gbIsFDBInstalled = False Then
        sSql = "select MEDID, DESCDISPLAY, DRUG_CHECK_DISABLED, DEFAULT_ROUTE " _
             & "from RX_LIST_FDB" _
             & " where"
         
        sSql = sSql & " ( upper(descdisplay) like " & goConn.SQLQuote("%" & UCase$(sSearchString) & "%") _
                & " OR upper(SYNONYMS) like " & goConn.SQLQuote("%" & UCase$(sSearchString) & "%") & " )"
             
        If IncludeInactive = False Then
            sSql = sSql & " and (DRUG_CHECK_DISABLED = 0 " _
                 & " or DRUG_CHECK_DISABLED is null) "
        End If
    Else
        sSql = "select b.MEDID, b.DESCDISPLAY, a.DRUG_CHECK_DISABLED, DEFAULT_ROUTE " _
             & "from  fdb.FDB_DISPENSABLE b " _
             & ", RX_LIST_FDB a" _
             & " where b.medid = a.medid  "
         
        sSql = sSql & " and ( upper(b.descdisplay) like " & goConn.SQLQuote("%" & UCase$(sSearchString) & "%") _
                & " OR upper(A.SYNONYMS) like " & goConn.SQLQuote("%" & UCase$(sSearchString) & "%") & " )"
             
        If IncludeInactive = False Then
            sSql = sSql & " and (a.DRUG_CHECK_DISABLED = 0 " _
                 & " or a.DRUG_CHECK_DISABLED is null) "
        End If
    End If
    
    Set objRS = goDAS.OpenRecordset(sSql)
    If Not (objRS.BOF And objRS.EOF) Then
        objRS.MoveFirst
        While (Not objRS.EOF)
            Set oMed = New Drug2Drug.MedData
            oMed.Description = objRS.Fields("DESCDISPLAY").Value & vbNullString
            oMed.ID = objRS.Fields("MEDID").Value & vbNullString
            oMed.IsActive = "0" & objRS.Fields("DRUG_CHECK_DISABLED").Value = 0
            oMed.Route = objRS.Fields("DEFAULT_ROUTE").Value & vbNullString
            oMeds.Add oMed
            objRS.MoveNext
        Wend
    End If
    
    Search_CHCSII_Meds = 0
    Exit Function
Err_Handler:
    Call gobjShared.ShowVBError(Err.Number, "Error Searching CHCSII FDB Formulary drug list", "CHCSII_FDB.Search_CHCSII_Meds", "DRUG INTERACTION SCREENING", vbExclamation)
    Search_CHCSII_Meds = -1
    Exit Function
    Resume
End Function

Public Function SearchMeds(ByRef oMeds As Collection, sSearchString As String) As Integer
    
    Dim oSearchFilter As Object ' SearchFilter
    Dim oDrugFilter As Object ' DrugSearchFilter
    Dim oFDBMeds As Object ' DispensableDrugs
    Dim oFdbMed As Object ' DispensableDrug
    Dim oMed As Object ' Drug2Drug.MedData

    On Error GoTo ErrHandler

    '*** Clear out previous Searched Result
    Set oMeds = Nothing
    Set oMeds = New Collection
    
    Call SetCreateDAS


    'Set oSearchFilter = New SearchFilter
    Set oSearchFilter = CreateObject("FDBDIF32.SearchFilter")
    oSearchFilter.SearchMethod = 2 'fdbSMExhaustive
    oSearchFilter.PhoneticSearch = 2 'fdbPSAfterEmpty
    
    'Set oDrugFilter = New DrugSearchFilter
    Set oDrugFilter = CreateObject("FDBDIF32.DrugSearchFilter")
    
    If sSearchString = "" Then
        MsgBxARMd "Please enter a search string.", vbOKOnly + vbInformation
        Exit Function
    Else
        Set oFDBMeds = goFDBNav.DispensableDrugSearch(sSearchString, oSearchFilter, oDrugFilter)
    End If
    
    For Each oFdbMed In oFDBMeds
        '*** Remove all INACTIVE drugs from the Search Result list
        '59830  Open  Inpt II Rx Checking-16 Update current Drug List
        If GetActiveStatus(oFdbMed.ID) <> NotInFDBFormulary Then
            Set oMed = New Drug2Drug.MedData
            
            oMed.Description = oFdbMed.Description
            oMed.ID = oFdbMed.ID
            oMed.Route = oFdbMed.Route
            oMed.RouteAbbrev = oFdbMed.RouteAbbrev
            oMed.Strength = oFdbMed.Strength
            oMed.StrengthUnit = oFdbMed.StrengthUnit
            oMed.IsActive = GetActiveStatus(oFdbMed.ID) = DrugIsActive.Active
            oMeds.Add oMed
        End If
    Next oFdbMed
  
Exit Function

ErrHandler:
    Select Case Err.Number
        Case -2147220702
            Call gobjShared.ShowVBError(Err.Number, "Search Text Required", "CHCSII_FDB.SearchMeds", "DRUG INTERACTION SCREENING", vbExclamation)

            SearchMeds = -2
        Case Else
            Err.Raise Err.Number, "FDB.SearchMeds", Err.Description
            SearchMeds = -1
    End Select
    
    Exit Function
    Resume
End Function

Private Function GetActiveStatus(MedID As String) As DrugIsActive
    
    Dim objRS  As Recordset
    Dim sSql As String
    
    On Error GoTo Err_Handler
    
    GetActiveStatus = NotInFDBFormulary
    
    sSql = "select medid, DRUG_CHECK_DISABLED " _
           & "From RX_LIST_FDB " _
           & "where medid = " & MedID
    
    Set objRS = goDAS.OpenRecordset(sSql)
    If Not (objRS.BOF And objRS.EOF) Then
        If "0" & objRS.Fields("DRUG_CHECK_DISABLED") = 0 Then
            GetActiveStatus = Active
        Else
            GetActiveStatus = InActive
        End If
    End If
    
    Exit Function
Err_Handler:
    Call gobjShared.ShowVBError(Err.Number, "Error obtaining Drug Active Status", "CHCSII_FDB.GetActiveStatus", "DRUG INTERACTION SCREENING", vbExclamation)
    Exit Function
    Resume
End Function


Public Sub ScreenDrug(sUser_Unit As String, oColDrug As Collection, ByRef UserActionText As String, ByRef OverrideReasonID As String, ByRef OverrideComment As String, ByRef WarningInfo As String, Optional SeverityLevel As DrugCheckSeverityLevel_Enum, Optional IsPregnant As Boolean = False)
    
    Dim oFDBMeds As Object ' ScreenDrugs
    Dim oFDBAllergies As Object ' ScreenAllergies
    
    Dim oFdbMed As Object ' ScreenDrug
    Dim oFDBAllergy As Object ' ScreenAllergy

    'Collection of all Results
    Dim oWarnings As colMedWarnings
    
    Dim oItem As MedData
    
    Dim ProspectiveDrugID As String
    Dim ProspectiveDrugDesc As String
    
    Dim sResult As String
    Dim i As Long
    
    Dim rsAllergies As ADODB.Recordset
  
    Dim oColAllergy As colAllergyResults
    
    Dim oColFDBTypes As colFDBType
    
    Dim oNonFDBAllergy As NonFDBAllergy
    Dim oColNonFDBAllergies As Collection
    
    On Error GoTo ErrHandler
    
'''    frmWaiting.Display
    
    'OA
    If gbIsFDBInstalled = False Then
        Exit Sub
    End If
    
    Call SetCreateDAS
    
    'Load collection data into FDB Objects
    'Set oFDBMeds = New ScreenDrugs
    Set oFDBMeds = CreateObject("FDBDIF32.ScreenDrugs")
    
    'Set oFDBAllergies = New ScreenAllergies
    Set oFDBAllergies = CreateObject("FDBDIF32.ScreenAllergies")
    
    'Meds to Compare
    If (oColDrug Is Nothing) Then
        Exit Sub
    Else
    
        WarningInfo = ""
        
        Call Debug_Start("Start [ScreenDrug]")

        For Each oItem In oColDrug
            'Set oFdbMed = New ScreenDrug
            Set oFdbMed = CreateObject("FDBDIF32.ScreenDrug")
            
            oFdbMed.UseConnection goFDBConn
            With oItem
                On Error Resume Next
                oFdbMed.Load .ID, 2 'fdbDTDispensableDrug
                If Err.Number < 0 Then
                    '** Skip drug not found in the Formulary
'                    Err.Raise Err.Number, Err.Description, "Drug not found in the FDB database."
                Else
                    oFdbMed.Description = .Description
                    oFdbMed.Prospective = .Prospective
'                    oFDBMed.Prospective = True
                
                    If .Prospective = True Then
                        ProspectiveDrugID = .ID
                        ProspectiveDrugDesc = .Description
                    End If
                    oFDBMeds.AddItem oFdbMed
                End If
            End With
        Next oItem
    End If
    
    On Error GoTo ErrHandler
   
    Set rsAllergies = New ADODB.Recordset
    Set rsAllergies = GetAllergies(sUser_Unit)
    
    Set oColNonFDBAllergies = New Collection
    
    If Not (rsAllergies.BOF And rsAllergies.EOF) Then
        rsAllergies.MoveFirst
        While Not rsAllergies.EOF
            'Set oFDBAllergy = New ScreenAllergy
            Set oFDBAllergy = CreateObject("FDBDIF32.ScreenAllergy")
            
            oFDBAllergy.UseConnection goFDBConn
            
            ' NEW (FDB): fdb_ID ALLERGY_DESC
            ' OLD (Pre-FDB): allergenncid BN_GNN
            If Len("" & rsAllergies.Fields("FDB_ID").Value) > 0 Then
                oFDBAllergy.Load rsAllergies.Fields("FDB_ID").Value, _
                                 rsAllergies.Fields("FDB_ALLERGY_TYPE").Value
                If Err.Number < 0 Then
                    '** Error loading FDB Allergen
                    '** FDB_ID/Allergy Type not correct
                Else
                    oFDBAllergy.Description = "" & rsAllergies.Fields("ALLERGY_DESC").Value
                    oFDBAllergy.Prospective = False
                        
                    oFDBAllergies.AddItem oFDBAllergy
                End If
            Else
                If rsAllergies.Fields("OLD_BN_GNN").Value & vbNullString = "NKA" Then
                    '*** 5/17/05 - Frank
                    '*** Do not "No Known Allergies" as an allergy
                Else
                Set oNonFDBAllergy = New NonFDBAllergy
                ' SCR 65398
                ' 2/17/05 - Frank Added in GetAllergies():  A.ALLERGENNCID, c.BN_GNN AS OLD_BN_GNN
                
                    oNonFDBAllergy.Description = "" & rsAllergies.Fields("OLD_BN_GNN").Value
                    oNonFDBAllergy.ID = "" & rsAllergies.Fields("AllergenNCID").Value
                oColNonFDBAllergies.Add oNonFDBAllergy
                End If
            End If
            rsAllergies.MoveNext
        Wend
    End If
    
    If oColNonFDBAllergies.Count > 0 Then
        Call ShowCHCS_Allergies(oColNonFDBAllergies)
        Set oColNonFDBAllergies = Nothing
    End If
    
    If Not (oColAllergy Is Nothing) Then
'        For Each oAllergy In oColAllergy
        For i = 1 To oColAllergy.Count
            'Set oFDBAllergy = New ScreenAllergy
            Set oFDBAllergy = CreateObject("FDBDIF32.ScreenAllergy")
            
            oFDBAllergy.UseConnection goFDBConn
            
            On Error Resume Next
            oFDBAllergy.Load oColAllergy.Item(i).PK, CInt(oColAllergy.Item(i).AllergyType)
            If Err.Number < 0 Then
                '** Skip drug not found in the Formulary
'                    Err.Raise Err.Number, Err.Description, "Allergy not found in the FDB database."
            Else
                oFDBAllergy.Description = oColAllergy.Item(i).DescriptionText
                oFDBAllergy.InternalID = oColAllergy.Item(i).PK
                oFDBAllergy.Prospective = oColAllergy.Item(i).AllergyType
            End If
            oFDBAllergies.AddItem oFDBAllergy
        Next i
    End If
    
    Set oWarnings = Nothing
    Set oWarnings = New colMedWarnings
    
    Dim iWarningCount As Integer
    
    If oFDBMeds.Count > 0 Then
    
        '** Check DUPLICATE Drug
        Call CheckDuplicateTherapy(oFDBMeds, oWarnings)
        iWarningCount = oWarnings.Count
        
        If iWarningCount > 0 Then WarningInfo = "Duplicate Drug Warnings"
        
        '** Check Drug to Drug Interaction
        Call CheckDrug2Drug(oFDBMeds, GetSeverityLevel(SeverityLevel), oWarnings)
         
        If oWarnings.Count > iWarningCount Then
            If Len(WarningInfo) > 0 Then WarningInfo = WarningInfo & "; "
            WarningInfo = WarningInfo & "Drug Interaction Warnings"
            iWarningCount = oWarnings.Count
        End If
        
        If IsPregnant Then
            'Check Drug Pregnancy Intererance
            Call CheckDrug2Preg(oFDBMeds, oWarnings)
        
            If oWarnings.Count > iWarningCount Then
                If Len(WarningInfo) > 0 Then WarningInfo = WarningInfo & "; "
                WarningInfo = WarningInfo & "Drug Pregnancy Warnings"
                iWarningCount = oWarnings.Count
            End If
        End If
        
        If oFDBAllergies.Count > 0 Then
            Call CheckDrug2Allergy(oFDBMeds, oFDBAllergies, oWarnings)
        
            If oWarnings.Count > iWarningCount Then
                If Len(WarningInfo) > 0 Then WarningInfo = WarningInfo & "; "
                WarningInfo = WarningInfo & "Drug Allergy Warnings"
                iWarningCount = oWarnings.Count
            End If
        End If
    End If
    
    If oWarnings.Count > 0 Then
        Call frmWarning.Display(oWarnings, ProspectiveDrugID, ProspectiveDrugDesc)
        
        If frmWarning.OverrideWarnings Is Nothing Then
            '** NO WARNINGS
            UserActionText = ""
            OverrideComment = ""
            OverrideReasonID = ""
        Else
            UserActionText = Trim$(frmWarning.UserActionText)
            OverrideComment = Trim$(frmWarning.OverrideComment)
            OverrideReasonID = Trim$(frmWarning.OverrideReasonID)
        End If
    
        Unload frmWarning
    End If
    
'''    Unload frmWaiting
    
    Call Debug_Stop("Stop [ScreenDrug]")

    Exit Sub

ErrHandler:
    Call Debug_Stop("Stop [ScreenDrug]")
'''    Unload frmWaiting
    Err.Raise Err.Number, "FDB.ScreenDrug", Err.Description
    Exit Sub
    Resume
End Sub



Public Sub ScreenDrug_01(oColDrug As Collection, oColAllergies As Collection, ByRef UserActionText As String, ByRef OverrideReasonID As String, ByRef OverrideComment As String, Optional SeverityLevel As DrugCheckSeverityLevel_Enum, Optional IsPregnant As Boolean = False)
    
    '*** Old way of doing ScreenDrug, testing from OE_Rx
    '*** Hard code allergy in frmAllergyData.frm

    Dim oFDBMeds As Object ' ScreenDrugs
    Dim oFDBAllergies As Object ' ScreenAllergies
    
    Dim oFdbMed As Object ' ScreenDrug
    Dim oFDBAllergy As Object ' ScreenAllergy

    'Collection of all Results
    Dim oWarnings As colMedWarnings
    
    Dim oItem As MedData
    Dim oAllergy As clsAllergyData
    
    Dim ProspectiveDrugID As String
    Dim ProspectiveDrugDesc As String
    
    Dim sResult As String
    Dim idx As Long
    

    On Error GoTo ErrHandler

    Call SetCreateDAS
    
    'Load collection data into FDB Objects
    'Set oFDBMeds = New ScreenDrugs
    Set oFDBMeds = CreateObject("FDBDIF32.ScreenDrugs")
    
    'Set oFDBAllergies = New ScreenAllergies
    Set oFDBAllergies = CreateObject("FDBDIF32.ScreenAllergies")
    
    'Meds to Compare
    If (oColDrug Is Nothing) Then
        Exit Sub
    Else
        
        For Each oItem In oColDrug
            'Set oFdbMed = New ScreenDrug
            Set oFdbMed = CreateObject("FDBDIF32.ScreenDrug")
            
            oFdbMed.UseConnection goFDBConn
            With oItem
                On Error Resume Next
                oFdbMed.Load .ID, 2 'fdbDTDispensableDrug
                If Err.Number < 0 Then
                    '** Skip drug not found in the Formulary
'                    Err.Raise Err.Number, Err.Description, "Drug not found in the FDB database."
                Else
                    oFdbMed.Description = .Description
                    oFdbMed.Prospective = .Prospective
                    If .Prospective = True Then
                        ProspectiveDrugID = .ID
                        ProspectiveDrugDesc = .Description
                    End If
                End If
            End With
            oFDBMeds.AddItem oFdbMed
        Next oItem
    End If
    
    'Existing Patient Allery
    If Not (oColAllergies Is Nothing) Then
        For Each oAllergy In oColAllergies
            Set oFDBAllergy = CreateObject("FDBDIF32.ScreenAllergy")
            oFDBAllergy.UseConnection goFDBConn
            With oAllergy
                On Error Resume Next
                oFDBAllergy.Load .ID, CInt(.AllergyType)
                If Err.Number < 0 Then
                    '** Skip drug not found in the Formulary
'                    Err.Raise Err.Number, Err.Description, "Allergy not found in the FDB database."
                Else
                    oFDBAllergy.Description = .Description
                    oFDBAllergy.InternalID = .InternalAllergyID
                    oFDBAllergy.Prospective = .Prospective
                End If
            End With
            oFDBAllergies.AddItem oFDBAllergy
        Next oAllergy
    End If
    
    Set oWarnings = Nothing
    Set oWarnings = New colMedWarnings
    
    If oFDBMeds.Count > 0 Then
        '** Check DUPLICATE Drug
        Call CheckDuplicateTherapy(oFDBMeds, oWarnings)
       
       '** Check Drug to Drug Interaction
       Call CheckDrug2Drug(oFDBMeds, GetSeverityLevel(SeverityLevel), oWarnings)
         
        If IsPregnant Then
            'Check Drug Pregnancy Intererance
            Call CheckDrug2Preg(oFDBMeds, oWarnings)
        End If
    
        If oFDBAllergies.Count > 0 Then
            Call CheckDrug2Allergy(oFDBMeds, oFDBAllergies, oWarnings)
        End If
    End If
    
    If oWarnings.Count > 0 Then
        Call frmWarning.Display(oWarnings, ProspectiveDrugID, ProspectiveDrugDesc)
        
        If frmWarning.OverrideWarnings Is Nothing Then
            '** NO WARNINGS
            UserActionText = ""
            OverrideComment = ""
            OverrideReasonID = ""
        Else
            UserActionText = Trim$(frmWarning.UserActionText)
            OverrideComment = Trim$(frmWarning.OverrideComment)
            OverrideReasonID = Trim$(frmWarning.OverrideReasonID)
        End If
    
        Unload frmWarning
    End If
    
    Exit Sub

ErrHandler:

    Err.Raise Err.Number, "FDB.ScreenDrug_01", Err.Description
    Exit Sub
    Resume
End Sub


Private Function GetSeverityLevel(SeverityLevel As DrugCheckSeverityLevel_Enum) As Integer ' fdbDDIMSeverityFilter
    
    '*** Default Severity Level to 1 - Contraindicated
    
    Dim fdbSeverityLevel As Integer ' fdbDDIMSeverityFilter
        
    Select Case SeverityLevel
        Case DrugCheckSeverityLevel_Contraindicated
            fdbSeverityLevel = 1 ' fdbDDIMSFContraindicated
        Case DrugCheckSeverityLevel_Severe
            fdbSeverityLevel = 2 ' fdbDDIMSFSevere
        Case DrugCheckSeverityLevel_Moderate
            fdbSeverityLevel = 3 ' fdbDDIMSFModerate
        Case Else   '*** DEFAULT
            fdbSeverityLevel = 3 ' fdbDDIMSFModerate
    End Select
        
    GetSeverityLevel = fdbSeverityLevel
    
End Function

Private Sub CheckDuplicateTherapy(oFDBMeds As Object, MedWarnings As colMedWarnings)
                                            'ScreenDrugs
    '*** MedWarnings is OUTPUT param
    
    Dim oScreen As Object ' Screening
    
    Dim oResults As Object ' DTScreenResults
    Dim oResult As Object ' DTScreenResult
    
    Dim bProspectiveDrugWarning As Boolean
    Dim oWarning As MedWarning
    
    Dim oDrugItems As Object ' DTScreenDrugItems
    Dim oDrugItem As Object ' DTScreenDrugItem
    Dim sDrugs As String
          
    Dim idx As Long
    Dim idxDup As Long

    On Error GoTo ErrHandler

    'Set oScreen = New Screening
    Set oScreen = CreateObject("FDBDIF32.Screening")
    
    oScreen.UseConnection goFDBConn
    
    '** Use this to eliminate Result that checked existing drug against other existing drug
    '** Only interested in checking Prospective Drug against Existing Medication
    oFDBMeds.Item(1).CustomID = oFDBMeds.Item(1).ID

    'Perform the screening.
    'The screening options are:
    '   Retrospective screening (screen all drugs).
    '   Use number of duplications allowed from the database.
    Set oResults = oScreen.DTScreen(oFDBMeds, True, 2)
                                                    'fdbDAFDBOnly
    If MedWarnings Is Nothing Then
        Set MedWarnings = New colMedWarnings
    End If
    
    'See if the screening produced any results/hits.
    If oResults.Count > 0 Then
      
        'Retrieve the items in the results collection object.
        For idx = 1 To oResults.Count
            Set oWarning = New MedWarning
            
            Set oResult = oResults.Item(idx)
            
            oWarning.DrugClass = oResult.ClassDescription
            
            'Create a list of the duplicated drugs reported by this result
            'Get the duplicated drugs collection object.
            Set oDrugItems = oResult.DrugItems
            sDrugs = ""
            bProspectiveDrugWarning = False
            For idxDup = 1 To oDrugItems.Count
                'Get the description of each drug in the collection.
                Set oDrugItem = oDrugItems.Item(idxDup)
                bProspectiveDrugWarning = bProspectiveDrugWarning Or oDrugItem.DrugCustomID = oFDBMeds.Item(1).CustomID
                sDrugs = sDrugs & oDrugItem.DrugDescription & ", "
            Next idxDup
            If Len(sDrugs) > 2 Then
                sDrugs = Left(sDrugs, Len(sDrugs) - 2)
            End If

            If bProspectiveDrugWarning = True Then
                '** Do not show Warnings of Interaction between existing medication
                '** only show Warning of Prospective Drug
                MedWarnings.Add "Duplicate Drug", oResult.ScreenMessage, "", sDrugs
            End If
        Next idx

  End If

Exit Sub

ErrHandler:
    Err.Raise Err.Number, "FDB.CheckDuplicateTherapy", Err.Description
    Exit Sub
 
End Sub

Private Sub CheckDrug2Drug(oFDBMeds As Object, SeverityLevel As Integer, MedWarnings As colMedWarnings)
                                        'ScreenDrugs        'fdbDDIMSeverityFilter
    Dim oScreen As Object 'Screening
    
    Dim oResults As Object 'DDIMScreenResults
    Dim oResult As Object 'DDIMScreenResult
    
    Dim oWarning As MedWarning
    
    Dim sDrugs As String
          
    Dim idx As Long
    Dim idxDup As Long

    On Error GoTo ErrHandler

    Set oScreen = CreateObject("FDBDIF32.Screening")
    oScreen.UseConnection goFDBConn
    
    GoTo Old_Way
    
    '*******************************************
    '****
    Dim oMeds As Object ' ScreenDrugs
    Dim oMed As Object ' ScreenDrug
    
    
    
    Dim i As Integer
    
    For i = 2 To oFDBMeds.Count
        
        Set oMeds = CreateObject("FDBDIF32.ScreenDrugs")
        oMeds.AddItem oFDBMeds.Item(1)
        oMeds.Item(1).CustomID = oMeds.Item(1).ID
        oMeds.AddItem oFDBMeds.Item(i)
        
        Set oResults = oScreen.DDIMScreen(oMeds, True, SeverityLevel, False, False)
        If oResults.Count = 0 Then
            Set oMeds = CreateObject("FDBDIF32.ScreenDrugs")
            oMeds.AddItem oFDBMeds.Item(i)
            oMeds.Item(1).CustomID = oMeds.Item(1).ID
            oMeds.AddItem oFDBMeds.Item(1)
            Set oResults = oScreen.DDIMScreen(oMeds, True, SeverityLevel, False, False)
        End If
        
         'See if the screening produced any results/hits.
    If oResults.Count > 0 Then
    
        'Retrieve the items in the results collection object.
        Set oWarning = New MedWarning
        
        Set oResult = oResults.Item(idx)
    
        'DRUG 2 DRUG Interaction
        sDrugs = ""
        For Each oResult In oResults
            If oResult.Drug1CustomID = oFDBMeds.Item(1).CustomID Then
                '** Do not show Warnings of Interaction between existing medication
                '** only show Warning of Prospective Drug
                MedWarnings.Add "Drug 2 Drug", oResult.ScreenMessage, "", oResult.Drug1Description & " and " & oResult.Drug2Description, oResult.MonographID, 2, GetSeverityDescription(oResult.SeverityLevelCode)
            End If                                                                                                                                            'fdbMTDDIM
        Next oResult
    
    End If
    
    
    Next i
    
    
Old_Way:
  
    '** Use this to eliminate Result that checked existing drug against other existing drug
    '** Only interested in checking Prospective Drug against Existing Medication
    oFDBMeds.Item(1).CustomID = oFDBMeds.Item(1).ID
    
'    Set oResults = oScreen.DDIMScreen(oFDBMeds, True, fdbDDIMSFSevere, False, False)
    'SCR 59828  Open  Inpt II Rx Checking-14 Severity Level Config
    'Function DDIMScreen(Drugs As ScreenDrugs, ProspectiveOnly As Boolean, SeverityFilter As fdbDDIMSeverityFilter, UseCustomInteractions As Boolean, LoadUserCategories As Boolean, [LoadCustomStrings As Boolean = False], [CustomStringCategory As String], [LoadCategorizedSeverities As Boolean = False], [CategorizedSeverityCategory As String]) as object 'DDIMScreenResults
    Set oResults = oScreen.DDIMScreen(oFDBMeds, True, SeverityLevel, False, False)
    Debug.Print "ODDIMS: " & Now & " (" & oResults.Count & ")"

    If MedWarnings Is Nothing Then
        Set MedWarnings = New colMedWarnings
    End If
  
    'See if the screening produced any results/hits.
    If oResults.Count > 0 Then
    
        'Retrieve the items in the results collection object.
        Set oWarning = New MedWarning
        
        Set oResult = oResults.Item(idx)
    
        'DRUG 2 DRUG Interaction
        sDrugs = ""
        For Each oResult In oResults
'            If oResult.Drug1CustomID = oFDBMeds.Item(1).CustomID Then
                '** Do not show Warnings of Interaction between existing medication
                '** only show Warning of Prospective Drug
                MedWarnings.Add "Drug 2 Drug", oResult.ScreenMessage, "", oResult.Drug1Description & " and " & oResult.Drug2Description, oResult.MonographID, 2, GetSeverityDescription(oResult.SeverityLevelCode)
'            End If                                                                                                                                         fdbMTDDIM
        Next oResult
    
    End If
    
Exit Sub

ErrHandler:
    Err.Raise Err.Number, "FDB.CheckDrug2Drug", Err.Description
    Exit Sub

End Sub

Private Function GetSeverityDescription(SeverityCode As String) As String

    Dim sDesc As String

    Select Case Trim$(SeverityCode)
        Case "1"
            sDesc = "Contraindicated"
        Case "2"
            sDesc = "Severe"
        Case "3"
            sDesc = "Moderate"
        Case Else
            sDesc = "Unknown"
    End Select
    
    GetSeverityDescription = sDesc
    
End Function
Private Sub CheckDrug2Allergy(oFDBMeds As Object, oFDBAllergies As Object, ByRef MedWarnings As colMedWarnings)
                                        'ScreenDrugs            'ScreenAllergies
    Dim oScreen As Object 'Screening
    
    Dim oResults As Object ' DAMScreenResults
    Dim oResult As Object ' DAMScreenResult
    
    Dim oWarning As MedWarning

    On Error GoTo ErrHandler

    Set oScreen = CreateObject("FDBDIF32.Screening")
    oScreen.UseConnection goFDBConn
    
    '** Use this to eliminate Result that checked existing drug against other existing drug
    '** Only interested in checking Prospective Drug against Existing Medication
    oFDBMeds.Item(1).CustomID = oFDBMeds.Item(1).ID
    Set oResults = oScreen.DAMScreen(oFDBMeds, oFDBAllergies, True)
  
'    Set oColWarnings = New MedWarnOverride
    If MedWarnings Is Nothing Then
        Set MedWarnings = New colMedWarnings
    End If
    
    'See if the screening produced any results/hits.
    If oResults.Count > 0 Then
          
        'Retrieve the items in the results collection object.
        Set oWarning = New MedWarning
        
        'DRUG 2 Allergy Interaction
        For Each oResult In oResults
            '** Do not show Warnings of Interaction between existing medication
            '** only show Warning of Prospective Drug
            If oResult.DrugCustomID = oFDBMeds.Item(1).CustomID Then
                MedWarnings.Add "Drug Allergy", oResult.DrugDescription & " and " & oResult.AllergenDescription, "", oResult.ScreenMessage
            End If
        Next oResult
    End If
    
Exit Sub

ErrHandler:
    Err.Raise Err.Number, "FDB.CheckDrug2Allergy", Err.Description
    Exit Sub

End Sub

Private Sub CheckDrug2Preg(oFDBMeds As Object, MedWarnings As colMedWarnings)
                                    'ScreenDrugs
    Dim oScreen As Object 'Screening
    
    Dim oResults As Object ' PREGScreenResults
    Dim oResult As Object ' PREGScreenResult
    
    Dim oWarning As MedWarning

    On Error GoTo ErrHandler

    Set oScreen = CreateObject("FDBDIF32.Screening")
    oScreen.UseConnection goFDBConn
    
    '** Use this to eliminate Result that checked existing drug against other existing drug
    '** Only interested in checking Prospective Drug against Existing Medication
    oFDBMeds.Item(1).CustomID = oFDBMeds.Item(1).ID
    
    Set oResults = oScreen.PREGScreen(oFDBMeds, True, 3)
                                                    'fdbPREGSFWarning
    
    If MedWarnings Is Nothing Then
        Set MedWarnings = New colMedWarnings
    End If
    
    'See if the screening produced any results/hits.
    If oResults.Count > 0 Then
      
        'Retrieve the items in the results collection object.
        Set oWarning = New MedWarning
        
        'DRUG Pregnancy Interaction
        For Each oResult In oResults
            '** Do not show Warnings of Interaction between existing medication
            '** only show Warning of Prospective Drug
            If oResult.DrugCustomID = oFDBMeds.Item(1).CustomID Then
                MedWarnings.Add "Drug Pregnancy", oResult.ScreenMessage, "", oResult.Description
            End If
        Next oResult
    End If
    
    
Exit Sub

ErrHandler:
    Err.Raise Err.Number, "FDB.CheckDrug2Preg", Err.Description
    Exit Sub

End Sub



Public Function ShowCHCSIIMeds(oColNonFDBDrugs As Collection) As Boolean

    On Error GoTo ErrHandler
  
    ShowCHCSIIMeds = True
    
    Call SetCreateDAS


    If oColNonFDBDrugs.Count = 0 Then Exit Function
    
    Load frmExistingMeds
    
    frmExistingMeds.LoadGrid oColNonFDBDrugs
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_OE_MED_APPLICATION, False
    frmExistingMeds.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_OE_MED_APPLICATION, False

Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    Resume
End Function

Public Function ShowCHCS_Allergies(oColNonFDBAllergies As Collection) As Boolean

    On Error GoTo ErrHandler
  
    ShowCHCS_Allergies = True
    
    Call SetCreateDAS


    If oColNonFDBAllergies.Count = 0 Then Exit Function
    
    Load frmExistingAllergies
    
    frmExistingAllergies.LoadGrid oColNonFDBAllergies
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_OE_MED_APPLICATION, False
    Screen.MousePointer = vbDefault
    frmExistingAllergies.Show vbModal
    Screen.MousePointer = vbHourglass
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_OE_MED_APPLICATION, False


Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetMonographID(sMedId As String, sMedName As String) As Long

    Dim objMono As Object 'Monograph
    Dim oDrug As Object 'DispensableDrug
    
    Call SetCreateDAS
    
    Set oDrug = CreateObject("FDBDIF32.DispensableDrug")

    oDrug.UseConnection goFDBConn
      
    oDrug.Load sMedId, 3, sMedName
                    'fdbDDLTMedID
                    
    Set objMono = oDrug.GetPEMMonograph("FDB-C", 1)
                                                'fdbMSFDBOnly

    GetMonographID = objMono.ID
    
Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ShowDrugInfo(sMedId As String, sMedName As String)


    '*********************
    '** Use MicroMedex instead of FDB
    Call DisplayMedex(sMedName)
    
    
    Exit Sub
    
    '***********************************
    '** FDB

      
''    Dim oDrug as object 'DispensableDrug
''    Dim objMono as object 'Monograph
''
''
''    Call SetCreateDAS
''
''    Set oDrug = CreateObject("FDBDIF32.DispensableDrug")
''
''    oDrug.UseConnection goFDBConn
''
''    oDrug.Load sMedId, fdbDDLTMedID, sMedName
''
''    Set objMono = oDrug.GetPEMMonograph("FDB-C", fdbMSFDBOnly)
''
''    If Not objMono Is Nothing Then
''        Load frmMonograph
''        frmMonograph.LoadMonograph objMono
''        '*** Ignore error can't show modal form from a modal form
''        On Error Resume Next
''
''        '12/23/2004 Frank Lourng
''        cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_OE_MED_APPLICATION, False
''        frmMonograph.Show
''        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_OE_MED_APPLICATION, False
''
''
''    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function GetPEMMonograph(sMedId As String, sMedName As String) As String

    Dim sMonograph As String
    
    Dim oDrug As Object 'DispensableDrug
    Dim objMono As Object 'Monograph
    
    On Error GoTo Err_Handler
    
    GetPEMMonograph = ""
    
    Call SetCreateDAS

    If sMedId = vbNullString Then
        sMedId = getMedID(sMedName)
    End If
    
    If sMedName = vbNullString Then
        sMedName = getMedDesc(CStr(sMedId))
    End If

    Set oDrug = CreateObject("FDBDIF32.DispensableDrug")

    oDrug.UseConnection goFDBConn

    Err.Clear
    On Error Resume Next
    oDrug.Load sMedId, 3, sMedName
                        'fdbDDLTMedID
    If Err.Number > 0 Then
        On Error Resume Next
        oDrug.Load sMedId, 6, sMedName
                        'fdbDDLTGCNSeqNo
    End If
    
    Set objMono = oDrug.GetPEMMonograph("FDB-C", 1)
                                            'fdbMSFDBOnly
    
    If objMono Is Nothing Then
        sMonograph = "Drug information is NOT available for this medication."
    Else
        sMonograph = objMono.GetFormattedText(1)
                                            'fdbMFPlainText
    End If
    
    
    GetPEMMonograph = sMonograph
        
    Exit Function
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

''Public Sub GetMed()
''
''    '************************
''    '** If MedID is not provided, the form will display MED Search option
''
''    Call SetCreateDAS
''
''    With frmDrugInfo
''        .Display
''        mMedID = .MedID
''        mMedDesc = .MedDesc
''        mNDC = .NDC
''        Call GetDefaultFDBAttributes(mMedID)
''    End With
''
''    Unload frmDrugInfo
''
''End Sub


Public Function ShowMonograph(MonographID As Long) As Boolean
 
    Dim oMonograph As Object 'Monograph

    Call SetCreateDAS


    Set oMonograph = CreateObject("FDBDIF32.Monograph")
    oMonograph.UseConnection goFDBConn
    oMonograph.Load 2, "FDB-C", MonographID, 1 ' fdbMSFDBOnly
                'fdbMTDDIM
    If Not oMonograph Is Nothing Then
        Load frmMonograph
        frmMonograph.LoadMonograph oMonograph
        frmMonograph.Show
    End If
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Sub DisplayPEMMonograph(Optional MedID As Variant, Optional MedDesc As Variant)

    '************************
    '** If MedID is not provided, the form will display MED Search option

    If gbIsFDBInstalled = False Then
        frmFDBCompat.Display True, "This functionality is provided by First Data Bank's NDDF Plus product if installed."
        Exit Sub
    End If
    
    Call SetCreateDAS


    If IsMissing(MedID) = True Then
        With frmPEMMonograph
            .Display
            mMedID = .MedID
            mMedDesc = .MedDesc
            mNDC = .NDC
            Call GetDefaultFDBAttributes(mMedID)
        End With
    Else
        frmPEMMonograph.Display MedID, MedDesc
    End If

    Unload frmPEMMonograph
    
End Sub

Private Sub GetDefaultFDBAttributes(MedID As String)

    Dim sSql As String
    Dim sTableName As String
    
    Dim oRs As ADODB.Recordset
    
    mDefaultRoute = ""
    mDefaultForm = ""
    mDefaultDose = ""
    mDosageStrength = ""
    mContentUnit = ""
    mNDC = ""

    If MedID = "" Or MedID = "0" Then
        Exit Sub
    End If
    
    Call SetCreateDAS
    
    sSql = "Select A.MEDID, A.STRENGTH, A.STRENGTHUNITS, B.ABBREV as DOSAGE_FORM, c.abbrev as DEFAULT_ROUTE " _
        & "From FDB_DISPENSABLE A, FDB_DOSEFORM B, FDB_ROUTE C " _
        & "Where A.DFID = B.DFID (+) " _
        & " and a.medid = " & MedID _
        & " and c.rtid = A.rtid "
    
    Set oRs = goDAS.OpenRecordset(sSql)
    
    '** Get Default Route/Default Dose
    If Not (oRs.EOF And oRs.EOF) Then
        oRs.MoveFirst
        mDefaultRoute = oRs.Fields("DEFAULT_ROUTE").Value & vbNullString
        mDefaultForm = oRs.Fields("DOSAGE_FORM").Value & vbNullString
        mDefaultDose = ""
        mDosageStrength = oRs.Fields("STRENGTH").Value & vbNullString
        mContentUnit = oRs.Fields("STRENGTHUNITS").Value & vbNullString
    End If
    
    '** Get Default NDC
    sSql = "Select distinct NDC from FDB_PACKAGEDDRUG " _
     & " where medid = " & MedID
    
    Set oRs = goDAS.OpenRecordset(sSql)
    
    If Not (oRs.EOF And oRs.EOF) Then
        oRs.MoveFirst
        mNDC = oRs.Fields("NDC").Value & vbNullString
    End If
    
    Set oRs = Nothing
    Exit Sub
End Sub


Private Function GetPatientAllergyIEN(sUser_Unit As String) As ADODB.Recordset

    Dim sSql As String
    
    On Error GoTo Err_Handler
     
    Call SetCreateDAS

    sSql = "select IEN from chcs_allergies "
    sSql = sSql & " where ien in (select allergenncid from Allergies"
    sSql = sSql & " where unit_number = " & goConn.SQLQuote(sUser_Unit) & ")"
 
    sSql = UCase$(sSql)
    Set GetPatientAllergyIEN = goDAS.OpenRecordset(sSql)
    Exit Function
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function GetPatientMedOrders(PatientUnitNumber As String) As ADODB.Recordset

    Dim sSql As String
    
    On Error GoTo Err_Handler
     
    Call SetCreateDAS

    sSql = "select * from medications "
    sSql = sSql & " where unit_number = " & goConn.SQLQuote(PatientUnitNumber)
 
    sSql = UCase$(sSql)
    Set GetPatientMedOrders = goDAS.OpenRecordset(sSql)
    Exit Function
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description


End Function



Private Function GetAllergies(sUser_Unit As String) As ADODB.Recordset

    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo Err_Handler
     
    Set rs = New ADODB.Recordset
    
    Call SetCreateDAS

''    sSql = "select  FDB_ID, FDB_ALLERGY_TYPE, ALLERGY_DESC from chcsii_fdb_allergies "
''    sSql = sSql & " where CHCS_ALLERGIES_IEN in (select allergenncid from Allergies"
''    sSql = sSql & " where unit_number = " & goConn.SQLQuote(sUser_Unit) & ")"
    
    '2/1/05 - modified to work with Allergy Module
    ' NEW (FDB): fdb_ID ALLERGY_DESC
    ' OLD (Pre-FDB): allergenncid BN_GNN
    ' 2/17/05 - Added:  A.ALLERGENNCID, c.BN_GNN AS OLD_BN_GNN
    '** 3/2/05 - Added: AND a.UPDATEFLAG = 'N' to exclude deleted rows
    sSql = sSql & "SELECT b.id, b.fdb_ID, b.FDB_ALLERGY_TYPE, "
    sSql = sSql & " b.ALLERGY_DESC, c.BN_GNN,  A.ALLERGENNCID, c.BN_GNN AS OLD_BN_GNN"
    sSql = sSql & " from allergies a LEFT JOIN ALLERGIES_FDB_LIST b ON a.fdb_seq_id = b.ID LEFT JOIN CHCS_Allergies c on a.ALLERGENNCID = c.ien"
    sSql = sSql & " where a.unit_number = " & goConn.SQLQuote(sUser_Unit)
    sSql = sSql & " AND a.UPDATEFLAG = 'N'"
    
    sSql = UCase$(Trim$(sSql))
    Set rs = goDAS.OpenRecordset(sSql)
    
    Set GetAllergies = rs
    
    Exit Function
Err_Handler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetRxExpandedSig(RxSig As String, Qty As String, Refill As String, Optional expandSIG As Integer = 1) As String

    '*** 4/1/05 -- Frank
    '*** Moved from GEM_OE.  This need to be call from Inpatient Order Management also

    '*** SCR 65972
    '*** 2/17/05 - Frank
    '*** Move from OE_SessionManager.PerformAction
    '*** called directly from Frm_OE_Rx
    '*** When Quantity or Refill is changed
    '*** When User enter a space bar on the Sig text box
    '*** when User click on the Exp. Sig button
    '*** Also call from OE_SessionManager.PerformAction
        
    '*** To resolve problem of when user Click Submit button
    '*** It does the exp. sig and stop.  User had to click Submit button
    '*** a second time.
    
    ' SCR 102777, JQ, 8/8/06 - added optional 4th arg to allow SIG expansion to be bypassed
    
    Dim sSql As String
    
    Dim tempSigCode As String
    Dim aSettings() As String
    Dim i As Integer
    Dim expSigCode As String
    Dim bSkipLookup As Boolean
    Dim sSkipflag As String  ' last 3 char of the SIG
            
    On Error GoTo ErrHandler
            
    Dim mobjdas As ICHCSII_DAS
    Dim I_Das As CHCSII_CONN.Conn
    Dim rs As ADODB.Recordset
            
            
    Set I_Das = New CHCSII_CONN.Conn
    Set mobjdas = I_Das.CHCSII_DAS
    
    tempSigCode = UCase$(Trim$(RxSig))
    expSigCode = vbNullString
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    ' SCR 102777, JQ, 8/8/06 -- allow user to toggle SIG expansion on and off
    bSkipLookup = (expandSIG <> 1)  ' passed in
    If tempSigCode <> "" Then
        sSkipflag = Right$(tempSigCode, 3)
    End If
    If sSkipflag = "||0" Then
        bSkipLookup = True
        tempSigCode = Left$(tempSigCode, Len(tempSigCode) - 3)  ' remove the flag
        ' lookup will be skipped
    ElseIf sSkipflag = "||1" Then
        bSkipLookup = False ' explicit (not really using this...it's the default)
        tempSigCode = Left$(tempSigCode, Len(tempSigCode) - 3)  ' remove the flag
        ' continue with lookup
    End If
    
    If bSkipLookup Then
        expSigCode = tempSigCode ' pass-thru
    Else
        ' this section does a word-by-word lookup of potential SIG abbreviations
        aSettings = Split(tempSigCode, " ")
        For i = LBound(aSettings) To UBound(aSettings)
        
            Dim hasDecimal As Boolean
            hasDecimal = (InStr(aSettings(i), "."))
            
            If hasDecimal Then
                Dim valuePresent As Boolean
                Dim beforeDecimal As String
                Dim afterDecimal As String
                
                valuePresent = (InStr(aSettings(i), ".") - 1)
                If valuePresent Then
                    beforeDecimal = mID(aSettings(i), InStr(aSettings(i), ".") - 1, (InStr(aSettings(i), ".") - 1))
                Else
                    'Nothing before the decimal so force IsNumeric to false...needed for ".5" values
                    beforeDecimal = "x"
                End If
                
                valuePresent = (InStr(aSettings(i), ".") + 1)
                If valuePresent Then
                    afterDecimal = mID(aSettings(i), InStr(aSettings(i), ".") + 1, (InStr(aSettings(i), ".") + 1))
                Else
                    'Nothing after the decimal so force IsNumeric to false
                    afterDecimal = "x"
                End If
                
                If Not IsNumeric(beforeDecimal) And IsNumeric(afterDecimal) Then
                    aSettings(i) = Replace(aSettings(i), ".", "0.")
                End If
                
                valuePresent = False
                beforeDecimal = vbNullString
                afterDecimal = vbNullString
                hasDecimal = False
            End If
            
            If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
                
            '** 2/4/05 - Frank
            '** Use Rx_Code_Rx which contains all Abbreviation (Frequency/Route/...)
            sSql = "select codedesc from rx_sig_code " & _
                    " where abbrev = " & I_Das.SQLQuote(aSettings(i))
                    
            Set rs = mobjdas.OpenRecordset(sSql)
            If Len(expSigCode) > 0 Then expSigCode = expSigCode & " "
            If Not rs.EOF Then
                ' we found a lookup
                expSigCode = expSigCode & rs("codedesc")
                aSettings(i) = vbNullString  ' because we're replacing it with its lookup value
            Else
                '** Do not add Qty or Refill back to SIG
                '** Already done in LostFocus()
                Dim bSkip As Boolean
                
                bSkip = InStr(1, aSettings(i), "#") > 0
                
                '*** 5/26/05 - Frank
                '*** Remove X = Refills
                '*** Add X in Rx_Sig_Code to means "For"
''                bSkip = bSkip Or InStr(1, aSettings(i), "X") > 0
                ''INSERT INTO RX_SIG_CODE VALUES(0,3, 'X', 'for')
                
                'bSkip = bSkip Or InStr(1, aSettings(i), "RFX") > 0
                'bSkip = bSkip Or InStr(1, aSettings(i), "RF") > 0
                
                'SCR 107029, JQ, 6/7/07
                If InStr(1, aSettings(i), "RF") > 0 Then
                    ' grab incoming #refills from the unexpanded SIG
                    Dim j As Integer
                    Dim originalRefills As String
                    j = 1
                    While j < Len(aSettings(i)) And mID$(aSettings(i), j, 1) Like "[!0-9]"
                        j = j + 1
                    Wend
                    If (j <= Len(aSettings(i))) Then
                        originalRefills = mID$(aSettings(i), j)
                    Else
                        originalRefills = "0"
                    End If
                    Refill = originalRefills
                    bSkip = True
                End If
                
                If bSkip = False Then
                    expSigCode = expSigCode & aSettings(i)
                End If
            End If
        Else
            sSql = "select RX_FREQUENCY_DESCRIPTION from rx_frequency " & _
                    " where RX_FREQUENCY_CODE = " & I_Das.SQLQuote(aSettings(i)) 'SCR 50608 -Shen 3/2/2004
            
            Set rs = mobjdas.OpenRecordset(sSql)
            
            If Not rs.EOF Then
                expSigCode = expSigCode & _
                    rs("RX_FREQUENCY_DESCRIPTION") & ", "
                aSettings(i) = vbNullString
                
            Else ' check route table
                sSql = "select RX_ROUTE_DESCRIPTION from rx_route " & _
                       "where RX_ROUTE_CODE = " & I_Das.SQLQuote(aSettings(i)) 'SCR 50608 -Shen 3/5/2004
        
            End If
                    
            Set rs = mobjdas.OpenRecordset(sSql)
            If Not rs.EOF Then
                expSigCode = expSigCode & _
                    rs(0) & ", "
                aSettings(i) = vbNullString
            End If
        End If
    Next
    End If ' if expanding
    
    expSigCode = Trim$(expSigCode)
    If Right$(expSigCode, 1) = "," Then
        expSigCode = Left$(expSigCode, Len(expSigCode) - 1)
    End If
    
    ' SCR 102777, JQ, 8/11/06 - if not expanding, don't tack on qty and refills either
    If Len(gobjShared.CmdLineSwitches("INPATIENT2")) > 0 And Not bSkipLookup Then
        If Len(expSigCode) > 0 Then expSigCode = expSigCode & "; "
        expSigCode = expSigCode & "Quantity " & Val(Qty)
        expSigCode = expSigCode & ", Refill " & Val(Refill)
    Else
        '*** Append elsewhere...
    End If
    
    GetRxExpandedSig = Replace$(expSigCode, "  ", " ")
    Set mobjdas = Nothing

    Exit Function
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CHCSII_FDB.GetExpanedSig", "DRUG2DRUG", vbExclamation)
    Exit Function
    Resume
    
End Function




Public Function preExpSig(ByRef txtQuantity As Object, _
                          ByRef txtRefills As Object, _
                          SIG As String, _
                          Optional TrimText As Boolean = True, _
                          Optional replaceRefills As Boolean = False)
    
    '*** SCR 66593
    '*** 2/23/05 - Frank
    '*** Add PreExpSig processing for translating Quantity and Refill
    
    Dim i As Integer
    Dim sRefill As String
    Dim sNewSig As String
    Dim arrSig() As String
    Dim bRefill As Boolean
    Dim sCode As String
    Dim bNotQtyRefill As Boolean
    
    Dim sChar As String
    Dim k As Integer
    
    Dim iNonNumericPos As Integer
    Dim sNonNumeric As String
    
    Dim tempRefills As String
    
'    If txtQuantity.Text = "" Then txtQuantity.Text = 0
'    If txtRefills.Text = "" Then txtRefills.Text = 0
'
'
'    '** Qty: Add #nn to SIG if missing
'    If InStr(1, SIG, "#") = 0 Then
'        k = InStr(1, SIG, "RF")
'        If k = 0 Then
'            SIG = SIG & " #" & txtQuantity.Text
'        Else
'            SIG = SIG & " "
'            tempRefills = mID$(SIG, k, InStr(k, SIG, " ") - k)
'            SIG = Left$(SIG, k - 1) & "#" & txtQuantity.Text & " " & tempRefills
'         End If
'    End If
'
'    '** Refills: Add RFnn to SIG if missing
'    If InStr(1, SIG, "RF") = 0 Then
'        If InStr(1, SIG, "RFX") = 0 Then
'            SIG = SIG & " RF" & txtRefills.Text
'        End If
'    End If
    
    arrSig = Split(SIG, " ")
    sNewSig = ""
    bRefill = False
    
    i = 0
    While i <= UBound(arrSig)
        bNotQtyRefill = True
        sCode = UCase$(Trim$(arrSig(i)))

''        bRefill = Left$(sCode, 2) = "RF" _
''            Or Left$(sCode, 3) = "RFX" _
''            Or Left$(sCode, 1) = "X"

        bRefill = Left$(sCode, 2) = "RF" Or Left$(sCode, 3) = "RFX"
        
        If Left$(sCode, 1) = "#" Then
            '** Qty
            txtQuantity.Text = mID$(arrSig(i), 2)
            Call AllowOnlyNumeric(txtQuantity, sNonNumeric, iNonNumericPos)

            
            bNotQtyRefill = False
        ElseIf bRefill = True Then
            '** Refill Qty
            If i = UBound(arrSig) Then
                sRefill = GetRefill(sCode, "NONE")
            Else
                sRefill = GetRefill(sCode, arrSig(i + 1))
            End If
            
            ' SCR 107029, JQ, 6/8/07; don't use SIG refills; commented out the next line
            ' SCR 133731, JQ, 6/29/07; when coming from A/P Txt_RxSig_LostFocus, we need to do this:
            If replaceRefills Then
                txtRefills.Text = Val(sRefill)
            End If
            
            Call AllowOnlyNumeric(txtRefills, sNonNumeric, iNonNumericPos)
            If Val(txtRefills.Text) >= 0 Then
                bNotQtyRefill = False
            End If
        End If

        If bNotQtyRefill Then
            sNewSig = sNewSig & " " & arrSig(i)
        End If
        i = i + 1
    Wend

    
    '** Original
    If TrimText Then
        preExpSig = Trim$(SIG)
    Else
        preExpSig = SIG
    End If
    
End Function

Private Sub AllowOnlyNumeric(ByRef ctlTextBox As TextBox, ByRef NonNumeric As String, ByRef NonNumericPosition As Integer)

    Dim sQty As String
    Dim k As Integer
    Dim sChar As String
    
    Dim sOriginal As String
    
    sOriginal = ctlTextBox.Text
    NonNumericPosition = 0
    
    If Len(ctlTextBox.Text) > 4 Then ctlTextBox.Text = Left$(ctlTextBox.Text, 4)
    If Not IsNumeric(ctlTextBox.Text) Then
        sQty = ""
        
        '*** Ignore non-numeric characters
        For k = 1 To Len(ctlTextBox.Text)
            sChar = mID$(ctlTextBox.Text, k, 1)
            If Asc(sChar) >= 48 And Asc(sChar) <= 57 Then
                sQty = sQty & sChar
            Else
                NonNumericPosition = k
                NonNumeric = mID$(sOriginal, k)
                Exit For
            End If
        Next k
        ctlTextBox.Text = sQty
    End If

    
End Sub

Public Function GetRefill(SIG As String, RefillQty As String) As String
    
    '*** SCR 66593
    '*** 2/23/05 - Frank
    '*** Add PreExpSig processing for translating Quantity and Refill
    
    If IsNumeric(RefillQty) Then
        SIG = SIG & RefillQty
    End If
    
    If Left$(SIG, 3) = "RFX" Then
        If IsNumeric(mID$(SIG, 4)) Then
            GetRefill = mID$(SIG, 4)
        End If
''    ElseIf Left$(SIG, 1) = "X" Then
''        If IsNumeric(mID$(SIG, 2)) Then
''            GetRefill = mID$(SIG, 2)
''        End If
    ElseIf Left$(SIG, 2) = "RF" Then
        If IsNumeric(mID$(SIG, 3)) Then
            GetRefill = mID$(SIG, 3)
        End If

    End If


End Function

Public Function Get_FDB_Version() As String
    
    '*** 6/8/05 -- Frank
    '*** Retrieve FDB Version from FDB_Version and format as following:
    '*** First Databank Framework v3.1 (Dec 06, 2004)
    
    Dim sSql As String
    Dim objRS  As Recordset

    
    On Error GoTo Err_Handler
    
    If gbIsFDBInstalled = False Then
        Get_FDB_Version = "NDDF Plus 3.2 is not installed."
        Exit Function
    End If
    
    '*** Clear out previous Searched Result
    Call SetCreateDAS
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    
    sSql = "select * from  FDB_VERSION order by VERSIONKEY asc "
         

    Set objRS = goDAS.OpenRecordset(sSql)
    If Not (objRS.BOF And objRS.EOF) Then
        objRS.MoveLast
        
        Dim sDate As String
        sDate = objRS.Fields("ISSUEDATE").Value & vbNullString
        Get_FDB_Version = "First Databank Framework v" & objRS.Fields("DBVERSION").Value & vbNullString & " (" & Format(mID$(sDate, 5, 2) & "/" & mID$(sDate, 7) & "/" & Left$(sDate, 4), "mmm dd, yyyy") & ")"
    End If
    
    Exit Function
Err_Handler:
    MsgBxARMd "Error retrieving FDB Versin from the database.", vbCritical, "Get FDB Version"
    Exit Function
    Resume
End Function

Public Sub SearchMedName(ByRef MedID As String, ByRef MedName As String, ByRef IsPregnant As Boolean, PatientUnitNumber As String, Optional ExcludeIVDrug As Boolean = True, Optional IncludeInactiveDrug As Boolean = True, Optional NewMedOrder As Boolean = False)
    
    Call frmSelectNewMed.Display(MedID, MedName, IsPregnant, PatientUnitNumber, ExcludeIVDrug, IncludeInactiveDrug, NewMedOrder)
    Unload frmSelectNewMed

End Sub
