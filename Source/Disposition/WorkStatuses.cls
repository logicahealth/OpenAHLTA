VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkStatuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"WorkStatus"
Attribute VB_Ext_KEY = "Member0" ,"WorkStatus"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mobjShared As CWShared
Private mobjDas As ICHCSII_DAS
Private mszUnitNumber As String
Private mszEncounterID As String
Private mszFacilityNCID As String
Private mszUserNCID As String

Const ModName = "WorkStatuses"



Public Function Save() As Boolean

    If SaveVIAMMI() = True Then
        If SaveVIASQL() = True Then '{{BOOKMARK Patient Needs to be fixed for this to work}}
            gobjPatient.LoadPatient gobjPatient.UnitNumber, ""
        
            gobjComm.Message cwiRESET_VISUAL_INDICATORS, "", NCID_CW_APPLICATION, NCID_DISPOSITION_APPLICATION
            Save = True
        End If
    End If
        
End Function
Private Function SaveVIAMMI() As Boolean
    Dim objWS As WorkStatus
    On Error GoTo errXIT
    Dim objMember As Member
    Dim szIsChecked As String
    Dim szNCID As String
    Dim objMMI As MmiDoc
    
    On Error GoTo errXIT
    
    'Retrieve the MMI Data
    '<SCR 36876
    Select Case True
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR)
            Set objMMI = New MmiDoc
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater)
            Set objMMI = New thr_MmiDoc
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeITT)
            Set objMMI = New thr_MmiDoc
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCHCSI_GUI)
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeUnknown)
    End Select
    '>SCR 36876
    Call objMMI.CreateMember("Patient")
    objMMI.Member.UnitNumber = mszUnitNumber
    If objMMI.Get(True) = False Then
        Err.Raise vbObjectError, ModName, "No Special Work Status patient data found in the MMI "
    End If
    
    'Get The DataSet
    Set objMember = objMMI.Member
    
    'Update/Create new members
    For Each objWS In Me
        Select Case objWS.Selected
        Case esselected
            szIsChecked = "Y"
        Case esremoved
            szIsChecked = "N"
        End Select
    
        Dim Found As Boolean
        
        objMember.GenericData.GotoFirst
        Do While Not objMember.GenericData.EOF
'           This is Wrong
'            objMember.GenericData = objWS.NCID
'            objMember.GenericData.Type = objWS.NCID
'
            If objMember.GenericData.Type = objWS.NCID Then
                Found = True
                Exit Do
            End If
            objMember.GenericData.GoToNext
        Loop
        
        If Found Then
            'reset
            Found = False
        Else
            'this adds a Type if it's not found
            objMember.GenericData = objWS.NCID
        End If
        
        objMember.GenericData.ValueText = szIsChecked
    Next

    objMMI.WriteUpdate
    objMMI.ReleaseLock
    
    SaveVIAMMI = True
    
    Exit Function
errXIT:
    
    SaveVIAMMI = False
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, ModName & ".SaveVIAMMI", App.Title, vbExclamation)

End Function
Private Function SaveVIASQL() As Boolean
    Dim objWS As WorkStatus
    Dim szSQL As String
    Dim szSQL2 As String
    Dim objDAS As ICHCSII_DAS
    Dim objSQL As ICHCSII_SQL.ISqlOps
    
    Dim C As Collection
    
    On Error GoTo errXIT
    
    'Establish Target
    Set objSQL = gobjCHCSIIConn.CHCSII_DAS(Auto)  '''<SCR 36876
    '<SCR 36876
    Select Case True
    Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR)
        Set C = New Collection
        
    Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater)
    Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeITT)
    Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCHCSI_GUI)
    Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeUnknown)
    End Select
    '>SCR 36876
    'Save SQL record
    For Each objWS In Me
        If objWS.Insert = eiInsert Then
            szSQL = "INSERT INTO WORK_STATUS ("
            szSQL = szSQL & "PATIENT_NCID, ENCOUNTER_NUMBER, WORK_STATUS_NCID, WORK_STATUS_SELECTED, WORK_STATUS_QUALIFIED, FACILITY_ID, PROVIDER_NCID"
            szSQL = szSQL & ") VALUES ("
            szSQL = szSQL & "'%UnitNumber', '%EncounterID','%NCID','%CHECKED', '%QUAL', '%FacilityNCID', '%UserNcid')"

        Else
            szSQL = "UPDATE WORK_STATUS SET "
            szSQL = szSQL & "WORK_STATUS_SELECTED = '%CHECKED', WORK_STATUS_QUALIFIED = '%QUAL', PROVIDER_NCID = '%UserNcid' "
            szSQL = szSQL & "WHERE "
            szSQL = szSQL & "PATIENT_NCID = '%UnitNumber' AND ENCOUNTER_NUMBER = '%EncounterID' AND "
            szSQL = szSQL & "FACILITY_ID = '%FacilityNCID' AND WORK_STATUS_NCID = '%NCID'"
        End If
    
        szSQL = Replace(szSQL, "%UnitNumber", mszUnitNumber)
        szSQL = Replace(szSQL, "%EncounterID", mszEncounterID)
        szSQL = Replace(szSQL, "%FacilityNCID", mszFacilityNCID)
        szSQL = Replace(szSQL, "%UserNcid", mszUserNCID)
        
        szSQL = Replace(szSQL, "%NCID", objWS.NCID)
        szSQL = Replace(szSQL, "%CHECKED", objWS.Selected)
        szSQL = Replace(szSQL, "%QUAL", objWS.Qualified)
        
        objWS.Insert = eiUpate
        '<SCR 36876
        Select Case True
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR)
            C.Add szSQL 'Performance can be enhanced by grouping the queries into one call
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater)
            Call objDAS.ExecuteSQL(szSQL)
          
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeITT)
            Call objDAS.ExecuteSQL(szSQL)
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCHCSI_GUI)
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeUnknown)
        End Select
        '>SCR 36876
    Next
    
    'Performance can be enhanced by grouping the queries into one call
    If mobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR) Then   '<SCR 36876
        Call objSQL.ExecuteMulti(C)
    End If
    
    SaveVIASQL = True
    
    Exit Function
errXIT:
    
    SaveVIASQL = False
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, ModName & ".SaveVIASQL", App.Title, vbExclamation)
End Function

Public Function Retrieve(ByVal UnitNumber As String, ByVal EncounterID As String, _
                         ByVal FacilityNCID As String, ByVal UserNCID As String) As Boolean

    On Error GoTo errXIT
    
    mszUnitNumber = UnitNumber
    mszEncounterID = EncounterID
    mszFacilityNCID = FacilityNCID
    mszUserNCID = UserNCID
    
    Call InitCollection
    
    'Attempt to get Data from Work Status table (if encounter is virgin, there will definitely be no data)
    'if no rows found, get the data from MMI
    If RetrieveVIASQL(UnitNumber, EncounterID, FacilityNCID) = False Then
        If RetrieveVIAMMI(UnitNumber) = False Then
            Err.Raise vbObjectError, ModName, "No Workstatus Record Was found"
        End If
    End If
    
    Retrieve = True
    
    Exit Function
    
errXIT:
    Retrieve = False
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, ModName & ".RetrieveVIASQL", App.Title, vbExclamation)
    
End Function
Private Function InitCollection()
    Set mCol = New Collection

    Add NCID_DIVING_STATUS, esremoved, eqUnspecified, DecodeNCID(NCID_DIVING_STATUS), eiInsert, NCID_DIVING_STATUS
    Add NCID_FLYING_STATUS, esremoved, eqUnspecified, DecodeNCID(NCID_FLYING_STATUS), eiInsert, NCID_FLYING_STATUS
    Add NCID_JUMPING_STATUS, esremoved, eqUnspecified, DecodeNCID(NCID_JUMPING_STATUS), eiInsert, NCID_JUMPING_STATUS
    Add NCID_MILITARY_POLICE, esremoved, eqUnspecified, DecodeNCID(NCID_MILITARY_POLICE), eiInsert, NCID_MILITARY_POLICE
    Add NCID_SUBMARINE, esremoved, eqUnspecified, DecodeNCID(NCID_SUBMARINE), eiInsert, NCID_SUBMARINE
    Add NCID_ON_MOBILITY, esremoved, eqUnspecified, DecodeNCID(NCID_ON_MOBILITY), eiInsert, NCID_ON_MOBILITY
    
End Function
'Retrieve Work_Status VIA the MMI
'for New Encounters that have not been saved.
'There may not be any MMI Work Status data, or the data may be incomplete
Private Function RetrieveVIAMMI(ByVal UnitNumber As String) As Boolean
    
    Dim objMember As Member
    Dim lIsChecked As EnumSelected
    Dim szNCID As String
    Dim objMMI As MmiDoc
    
    On Error GoTo errXIT
    
    'Retrieve the MMI Data
    '<SCR 36876
    Select Case True
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR)
            Set objMMI = New MmiDoc
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater)
            Set objMMI = New thr_MmiDoc
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeITT)
            Set objMMI = New thr_MmiDoc
            
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeCHCSI_GUI)
        Case mobjShared.IsAppMode(enumCSCSIIAppMode.modeUnknown)
    End Select
    '>SCR 36876
    Call objMMI.CreateMember("Patient")
    objMMI.Member.UnitNumber = UnitNumber
    If objMMI.Get(False) = False Then
        Err.Raise vbObjectError, ModName, "No Special Work Status patient data found in the MMI "
    End If
    
    'Walk The DataSet
    Set objMember = objMMI.Member
    objMember.GenericData.GotoFirst
    
    Do While Not objMember.GenericData.EOF
    
        
        szNCID = objMember.GenericData.Type
        If NCIDOfInterest(szNCID) Then
        
            Select Case objMember.GenericData.ValueText
                Case "Y"
                    lIsChecked = esselected
                Case "N"
                    lIsChecked = esremoved
            End Select
            
            'Me.Item(szNCID).Insert = eiInsert
            Me.Item(szNCID).Qualified = eqUnspecified
            Me.Item(szNCID).Qualified_Initial = eqUnspecified
            Me.Item(szNCID).Selected = lIsChecked
            Me.Item(szNCID).Selected_Initial = lIsChecked
            
        End If
                           
        objMember.GenericData.GoToNext
    Loop
    RetrieveVIAMMI = True
    
    Exit Function
    
errXIT:
    RetrieveVIAMMI = False
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, ModName & ".RetrieveVIAMMI", App.Title, vbExclamation)
 End Function

'If the Works_Status table has already been written to, retrieve it
Private Function RetrieveVIASQL(ByVal UnitNumber As String, ByVal EncounterID As String, ByVal FacilityNCID As String) As Boolean

    Dim rs    As Recordset
    Dim i     As Long
    Dim szSQL As String
    Dim szNCID As String
    Dim objWS As WorkStatus
    
    On Error GoTo errXIT
    
    'Set The Default Insert State
    For Each objWS In Me
        objWS.Insert = eiInsert
    Next
    
    szSQL = "SELECT WORK_STATUS_NCID, WORK_STATUS_SELECTED, WORK_STATUS_QUALIFIED " & vbLf _
         & "FROM WORK_STATUS " & vbLf _
         & "WHERE PATIENT_NCID = '%1'  AND ENCOUNTER_NUMBER = '%2'  AND FACILITY_ID = '%3'"
    szSQL = Replace(szSQL, "%1", UnitNumber)
    szSQL = Replace(szSQL, "%2", EncounterID)
    szSQL = Replace(szSQL, "%3", FacilityNCID)
    
    
    Set rs = mobjDas.OpenRecordset(szSQL)
    
    With rs
        
        If rs.RecordCount = 0 Then Exit Function
    
        Do While Not .EOF
        
            szNCID = .Fields("WORK_STATUS_NCID").Value
            If NCIDOfInterest(szNCID) Then
                Me.Item(szNCID).Insert = eiUpate    'Mark the Update state for use during saves
                Me.Item(szNCID).Qualified = .Fields("WORK_STATUS_QUALIFIED")
                Me.Item(szNCID).Qualified_Initial = .Fields("WORK_STATUS_QUALIFIED")
                Me.Item(szNCID).Selected = .Fields("WORK_STATUS_SELECTED")
                Me.Item(szNCID).Selected_Initial = .Fields("WORK_STATUS_SELECTED")
            End If
            .MoveNext
        Loop
    End With
    
    RetrieveVIASQL = True
    
    Exit Function
    
errXIT:
    RetrieveVIASQL = False
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, ModName & ".RetrieveVIASQL", App.Title, vbExclamation)
    
End Function
Public Function DecodeNCID(ByVal NCID As String) As String

    Select Case NCID
        Case NCID_PERSONAL_RELIABILTY_PROGRAM
            DecodeNCID = "PRP/SCI/PSP"
        Case NCID_JUMPING_STATUS
            DecodeNCID = "Jumping"
        Case NCID_DIVING_STATUS
            DecodeNCID = "Diving"
        Case NCID_MILITARY_POLICE
            DecodeNCID = "Military Police"
        Case NCID_FLYING_STATUS
            DecodeNCID = "Flying"
        Case NCID_SUBMARINE
            DecodeNCID = "Submarine"
        Case NCID_ON_MOBILITY
            DecodeNCID = "On Mobility"
        Case NCID_COMMAND_SECURITY
            DecodeNCID = "Command Security"
        
    End Select
    
End Function
Public Function NCIDOfInterest(ByVal NCID As String) As Boolean

    Select Case NCID
        Case NCID_JUMPING_STATUS
            NCIDOfInterest = True
        Case NCID_DIVING_STATUS
            NCIDOfInterest = True
        Case NCID_MILITARY_POLICE
            NCIDOfInterest = True
        Case NCID_FLYING_STATUS
            NCIDOfInterest = True
        Case NCID_SUBMARINE
            NCIDOfInterest = True
        Case NCID_ON_MOBILITY
            NCIDOfInterest = True
    End Select
    
End Function

Private Function Add(ByVal NCID As String, ByVal Selected As EnumSelected, ByVal Qualified As EnumQualified, _
                    ByVal DESCRIPTION As String, ByVal Insert As EnumInsert, Optional ByVal sKey As String) As WorkStatus
    'create a new object
    Dim objNewMember As WorkStatus
    Set objNewMember = New WorkStatus


    'set the properties passed into the method
    objNewMember.Insert = Insert
    objNewMember.Qualified = Qualified
    objNewMember.Qualified_Initial = Qualified
    objNewMember.Selected = Selected
    objNewMember.Selected_Initial = Selected
    objNewMember.DESCRIPTION = DESCRIPTION
    objNewMember.NCID = NCID
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As WorkStatus
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    count = mCol.count
End Property


Private Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
On Error GoTo ErrHnd
    'creates the collection when this class is created
    Set mCol = New Collection
    Set mobjShared = New CWShared
    Set mobjDas = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36876
    Exit Sub
ErrHnd:
    Resume Next
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mobjShared = Nothing
    Set mobjDas = Nothing
End Sub

