VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNonProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'If the user is a clinician we want them to be the first in the list, but only the first time
'The module is opened.
Private mboolFirstTime As Boolean
Private mboolSignableProvider As Boolean
Private lintEncounterProvider As Long
Private Const strName As String = "Non_Providers_OE"
Private Const strDataType As String = "Providers"
Private lngCollAdds As Long
'ADDProviders change
'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:09 PM
'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode
Private mobjConfig As Config



Public Property Get fldORDER() As String
fldORDER = "FLD_ORDER"
End Property
Public Property Get fldSignatureClass() As String
    fldSignatureClass = "SIGNATURE_CLASS"
End Property
Public Property Get fldProviderName() As String
    fldProviderName = "NAME"
End Property
Public Property Get fldIEN() As String
fldIEN = "IEN"
End Property
Public Property Get fldNCID() As String
fldNCID = "NCID"
End Property

Public Sub ClearOSH()
'*******************************
'*THIS IS A TEMPORARY FUNCTION *
'*FOR TESTING ONLY             *
'*******************************
'Dim strName As String
'Dim strDataType As String
Dim objOshMan As New OshMan
'strName = "Non_Providers_OE"
'strDataType = "Non_Providers_OE"
'If Not objOshMan.DeleteOshItems(strName, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID)) Then

If Not objOshMan.DeleteOshItems(strName, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID)) Then
        MsgBox "Cleanup of existing Providers  Failed.  Please see your system administrator.", vbCritical + vbOKOnly, "Non Provider Critical Error"
End If
End Sub

Public Property Get SignableProvider() As Boolean
SignableProvider = mboolSignableProvider
End Property
'Public Property Get EncounterProvider() As Integer
'EncounterProvider = mintEncounterProvider
'End Property
'Public Property Let EncounterProvider(intEncounterProvider As Integer)
'mintEncounterProvider = intEncounterProvider
'End Property
Private Sub ResetOSH()
'Dim strName As String
'Dim strDataType As String
Dim objOshMan As New OshMan
Dim iCount As Integer
Dim IntCollCount As Integer
'strName = "Non_Providers_OE"
'strDataType = "Providers"
'If there is more than 20 items remove the older items from the collection before writing
' to the OSH.  We don't want the list to grow out of control.
If gobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR 37894
    objOshMan.ConfigObject = mobjConfig
End If
If gobjColl.Count > 20 Then
    IntCollCount = gobjColl.Count
    For iCount = IntCollCount To 21 Step -1
        gobjColl.Remove iCount
    Next iCount
End If
'Call the Osman utility so that you can store the items.  First clear it and then rewrite.
If Not objOshMan.DeleteOshItems(strName, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID)) Then
        MsgBox "Cleanup of existing Providers  Failed.  Please see your system administrator.", vbCritical + vbOKOnly, "Non Provider Critical Error"
Else
    If Not objOshMan.CreateOshItems(strName, NCID_BIN_ITEM_TEXT_TYPE, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID), gobjColl) Then
        MsgBox "Providers save failed.", vbCritical + vbOKOnly, "Non Provider - Critical Error"
    End If
End If
Set objOshMan = Nothing

End Sub
Public Function SearchClinician(recProviders As ADODB.Recordset) As Boolean

Dim objClinician As Clinician
'Dim strName As String
'Dim strDataType As String
Dim objDas As ICHCSII_DAS
Dim objOshMan As New OshMan
Dim iCount As Integer
Dim vItem As Variant
Dim bFound As Boolean
Dim lngItemData As Long
'strName = "Non_Providers_OE"
'strDataType = "Non_Providers_OE"
Set objDas = SetCHCSII_DAS
Set objClinician = New Clinician

'Set Clinician to look for Only Providers (ClinicClient.vbp; ClinOps.cls modification)
objClinician.Type_Of_Search = Providers
'Since we modified ClinOps in the ClinicClient.vbp, we can search in the normal fashion
If objClinician.SearchClinician(CStr(glFacilityNCID), "", cwobjLogonEnviron, True, False) Then
    'put the new one at the top of the list so it will be the new number 1
    If gobjColl.Count = 0 Then
        gobjColl.Add objClinician.NCID, strName & " - " & gobjColl.Count + 1
        lngCollAdds = gobjColl.Count
    Else
        For Each vItem In gobjColl
            If vItem = objClinician.NCID Then
                bFound = True
                Exit For
            End If
        Next vItem
        If Not bFound Then
            'RFoster Begin:SCR #27699;   Developer: Robert Foster 09/18/2002 07:20 PM
            'RFoster Note: Function Changed: Search Clinician      Impact of Change:  Avoid Error 457
            gobjColl.Add objClinician.NCID, strName & " - " & lngCollAdds + 1, 1
            lngCollAdds = lngCollAdds + 1
        'RFoster End: SCR #27699;
        Else ' they searched for one already in their list...
            recProviders.MoveFirst
            recProviders.Find "NCID=" & objClinician.NCID
            lngItemData = recProviders.Fields("FLD_ORDER").Value
            Call ProviderSelected(recProviders, lngItemData)
            Exit Function
        End If
    End If
    'save the changes to the OSH based on the possible changes
    Call ResetOSH
    'if they changed something refetch the providers
    Call getProviders(recProviders)
    
End If




End Function
Public Function SearchAdditionalProviders(recProviders As ADODB.Recordset) As Boolean
'ADDProviders change
'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:09 PM
'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode
Dim objClinician As Clinician
'Dim strName As String
'Dim strDataType As String
Dim objDas As ICHCSII_DAS
Dim objOshMan As New OshMan
Dim iCount As Integer
Dim vItem As Variant
Dim bFound As Boolean
Dim lngItemData As Long
Set objDas = SetCHCSII_DAS
Set objClinician = New Clinician

'Set Clinician to look for all Providers (ClinicClient.vbp; ClinOps.cls modification)
objClinician.Type_Of_Search = NoFilter
'Since we modified ClinOps in the ClinicClient.vbp, we can search in the normal fashion
If objClinician.SearchClinician(CStr(glFacilityNCID), "", cwobjLogonEnviron, True, False) Then
    'put the new one at the top of the list so it will be the new number 1
    SearchAdditionalProviders = True

    If gobjColl.Count = 0 Then
        gobjColl.Add objClinician.NCID, strName & " - " & gobjColl.Count + 1
        lngCollAdds = gobjColl.Count
    Else
        For Each vItem In gobjColl
            If vItem = objClinician.NCID Then
                bFound = True
                Exit For
            End If
        Next vItem
        If Not bFound Then
            gobjColl.Add objClinician.NCID, strName & " - " & lngCollAdds + 1, 1
            lngCollAdds = lngCollAdds + 1
        Else ' they searched for one already in their list...
            recProviders.MoveFirst
            recProviders.Find "NCID=" & objClinician.NCID
            lngItemData = recProviders.Fields("FLD_ORDER").Value
            Call AdditionalProviderSelected(recProviders, lngItemData)
            Exit Function
        End If
    End If
    'save the changes to the OSH based on the possible changes
    Call ResetOSH
    'if they changed something refetch the providers
    Call getAdditionalProvidersList(recProviders)
    
End If



End Function
Public Property Let ConfigObject(objConfig As Object)
'ADDProviders change
'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:09 PM
'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode

Set mobjConfig = objConfig
End Property

Public Function getAdditionalProvidersList(recProviders As ADODB.Recordset) As Boolean
'ADDProviders change
'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:09 PM
'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode
Dim strSQL As String
Dim strTemp As String
Dim strFront As String
Dim strWhere As String
Dim iCount As Integer
Dim recTemp As ADODB.Recordset
Dim objOshMan As New OshMan
Dim objDas As ICHCSII_DAS
Set objDas = SetCHCSII_DAS
If gobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR 37894
    objOshMan.ConfigObject = mobjConfig
End If
Set gobjColl = objOshMan.GetOshItems(strName, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID))
'strFront = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "

strFront = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
strFront = strFront & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS " 'FROM PROVIDER "
strFront = strFront & " UNIT_NUMBER, SIGNATURE_CLASS FROM PROVIDER "
If gobjColl Is Nothing Then Set gobjColl = New Collection

If gobjColl.Count <> 0 Then ' Add the OSH items to the Select Statement
    strTemp = "WHERE NCID IN ("
    For iCount = 1 To gobjColl.Count
        If iCount = 1 Then
            strTemp = strTemp & gobjColl(iCount)
        Else
            strTemp = strTemp & ", " & gobjColl(iCount)
        End If
    Next iCount
    strTemp = strTemp & ")"
    If strTemp <> "WHERE NCID IN ()" Then
        strSQL = strFront & strTemp & " AND FACILITY_NCID = " & glFacilityNCID
    End If
    If recProviders.State <> adStateClosed Then
        recProviders.Close
    End If
End If
If strSQL <> "" Then
     Set recTemp = objDas.OpenRecordset(strSQL)
Else
    Set recTemp = New ADODB.Recordset
End If

BuildRecordset recTemp, recProviders




End Function

Public Function getProviders(recProviders As ADODB.Recordset) As Boolean
Dim objDas As ICHCSII_DAS
Dim recTemp As New ADODB.Recordset
Dim recBackUp As New ADODB.Recordset
Dim strCheck As String
Dim strFront As String
Dim strSQL As String
Dim strTemp As String
Dim iCount As Integer
'Dim strName As String
Dim boolNCIDFound As Boolean
Dim intTemp As Integer
Dim I As Integer
Dim intPosition As Integer
'strName = "Non_Providers_OE"
Set objDas = SetCHCSII_DAS
If mboolFirstTime = True Then ' If this is the first time through the routine then
                              ' set up the collection so that the provider is the first in the list.
    Select Case gintSig_Class
        Case 2, 3, 4 'They are a signable provider...
            mboolSignableProvider = True
            If gobjColl.Count = 0 Then ' Add them to the collection so that they will be in the OSH.
                gobjColl.Add glUserNCID, strName & " - " & gobjColl.Count + 1
                lngCollAdds = gobjColl.Count
            Else    'We need find them and to move them to the top
                For iCount = 1 To gobjColl.Count ' look for the provider and then move them to the top
                    If gobjColl(iCount) = glUserNCID Then
                        boolNCIDFound = True
                        ReOrderCollection CLng(iCount)
                    End If
                Next iCount
                If boolNCIDFound = False Then ' If the provider isn't found, then add them to the collection
                    gobjColl.Add glUserNCID, strName & " - " & gobjColl.Count + 1, 1
                    lngCollAdds = gobjColl.Count
                End If
            End If
'            Call ResetOSH
        Case Else ' If they're a non-provider or a outside provider then build the list from the
        mboolSignableProvider = False
    End Select    ' values that are stored in the OSH that were retrieved in OpenNonProvider...
    mboolFirstTime = False
End If
'Make sure that recTemp is closed
If recTemp.State <> adStateClosed Then
    recTemp.Close
End If
'Create the query to get the Providers
'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:57 AM
'R. Foster SCR #34505;   Developer: Robert Foster 04/24/2003 10:56 AM
'R. Foster Note: Function Changed:   CHCSII_NonProvider.dll\OpenNonProvider - OpenNonProvider; getProviders   CHCSII_Consult.dll\Form_Load      Impact of Change: Enable Saves to for NPOE in Provider GUI mode.
'Change for CHCS_GUI...
If gobjShared Is Nothing Then Set gobjShared = New CWShared
If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR 37894
    strFront = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
    'Robert Foster End: SCR #23369;
    If gobjColl.Count <> 0 Then ' Add the OSH items to the Select Statement
        strTemp = "WHERE CLINICIAN_NCID IN ("
        For iCount = 1 To gobjColl.Count
            If iCount = 1 Then
                strTemp = strTemp & gobjColl(iCount)
            Else
                strTemp = strTemp & ", " & gobjColl(iCount)
            End If
        Next iCount
        strTemp = strTemp & ")"
        If strTemp <> "WHERE CLINICIAN_NCID IN ()" Then
            'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:20 PM
            'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
            strSQL = strFront & strTemp & " AND FACILITY_NCID = " & glFacilityNCID
            'RFoster Begin:SCR #27658; 27570;   Developer: Robert Foster 09/17/2002 03:03 PM
            'RFoster Note: Function Changed: getProviders      Impact of Change: Prevent previously selected non-providers from getting into non-providers' provider list
            strSQL = strSQL & " AND SIGNATURE_CLASS IN(2,3,4)"
        'RFoster End: SCR #27658; 27570;
        'RFoster End: SCR #
        End If
        If recProviders.State <> adStateClosed Then
            recProviders.Close
        End If
        If strSQL <> "" Then
            Set recTemp = objDas.OpenRecordset(strSQL)
    '        now go through and make sure that we have them all...
            If recTemp.RecordCount <> gobjColl.Count Then 'some are missing; why am I looking into 2 tables?!?!?
                If recTemp.RecordCount > 0 Then 'We Have Some Records
                    For iCount = 1 To gobjColl.Count
                        strTemp = gobjColl(iCount) ' go through the collection and look for the record in the recordset
                        recTemp.MoveFirst
                        recTemp.Find "CLINICIAN_NCID=" & strTemp
                        If recTemp.EOF Then 'It ain't there we need to retrieve it from the Provider Table...
                            'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:57 AM
                            strSQL = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
                            'Robert Foster End: SCR #23369;
                            strSQL = strSQL & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
                            strSQL = strSQL & "WHERE NCID=" & gobjColl(iCount)
                            'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:20 PM
                            'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
                            strSQL = strSQL & " AND FACILITY_NCID = " & glFacilityNCID
                            'RFoster Begin:SCR #27658; 27570;   Developer: Robert Foster 09/17/2002 03:03 PM
                            'RFoster Note: Function Changed: getProviders      Impact of Change: Prevent previously selected non-providers from getting into non-providers' provider list
                            strSQL = strSQL & " AND SIGNATURE_CLASS IN(2,3,4)"
                            'RFoster End: SCR #27658; 27570;
                            'RFoster End: SCR #
                            'RFoster End: SCR #
                            Set recBackUp = Nothing
                            Set recBackUp = objDas.OpenRecordset(strSQL)
                            If recBackUp.RecordCount > 1 Then 'we have more than one uuuhhuhuhuhuh
                                intTemp = 0 'reset the checks
                                I = 0
                                Do Until recBackUp.EOF
                                    If Not IsNull(recBackUp.Fields("SIGNATURE_CLASS").Value) And Trim(recBackUp.Fields("SIGNATURE_CLASS").Value) <> "" Then
                                        If recBackUp.Fields("SIGNATURE_CLASS").Value > intTemp Then
                                            intTemp = recBackUp.Fields("SIGNATURE_CLASS").Value
                                            intPosition = I + 1 'set it so that you can move directly to the proper record later
                                        End If
                                    I = I + 1
                                    End If
                                    recBackUp.MoveNext
                                Loop
                                recBackUp.MoveFirst
                                recBackUp.Move intPosition 'Move to that record.
                                'Add what we got from the provider table
                                'To the information that we got from clinic_provider...
                                recTemp.AddNew
                                recTemp.Fields("PROVIDER_IEN").Value = CDbl(recBackUp.Fields("PROVIDER_IEN").Value)
                                recTemp.Fields("PROVIDER_NAME").Value = recBackUp.Fields("PROVIDER_NAME").Value
                                recTemp.Fields("CLINICIAN_NCID").Value = recBackUp.Fields("CLINICIAN_NCID").Value
                                recTemp.Fields("SIGNATURE_CLASS").Value = recBackUp.Fields("SIGNATURE_CLASS").Value
                                recTemp.Update
                            Else ' we only have one... or maybe 0...
                                If recBackUp.RecordCount = 1 Then 'IF we're here we should have something unless the Record
                                    recBackUp.MoveFirst           'was deleted. Just in case make sure we have a record
                                    recTemp.AddNew
                                    recTemp.Fields("PROVIDER_IEN").Value = CDbl(recBackUp.Fields("PROVIDER_IEN").Value)
                                    recTemp.Fields("PROVIDER_NAME").Value = recBackUp.Fields("PROVIDER_NAME").Value
                                    recTemp.Fields("CLINICIAN_NCID").Value = recBackUp.Fields("CLINICIAN_NCID").Value
                                    recTemp.Fields("SIGNATURE_CLASS").Value = recBackUp.Fields("SIGNATURE_CLASS").Value
                                    recTemp.Update
                                End If ' No Else for 0; The record was deleted and they can't choose it anyway
                            End If     ' I don't think you can delete a provider without causing a slew of other problems anyway
                        End If
                    Next iCount
                Else ' We have no records re-run the query against the Provider Table and be done with it...
                    'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:57 AM
                    strSQL = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
                    'Robert Foster End: SCR #23369;
                    strSQL = strSQL & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
                    'We can use the IN clause that we built above, but we need to modify it...
                    strTemp = Replace(strTemp, "CLINICIAN_NCID", "NCID")
                    strSQL = strSQL & strTemp 'strTemp was filled above...
                    'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:20 PM
                    'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
                    strSQL = strSQL & " AND FACILITY_NCID = " & glFacilityNCID
                    'RFoster End: SCR #
                    Set recTemp = Nothing
                    Set recTemp = objDas.OpenRecordset(strSQL) ' Now we should have the records that we're after.
                End If
            End If
            Set recBackUp = Nothing
            BuildRecordset recTemp, recProviders
        End If
        
    Else 'This is a non-provider and they have never chosen any providers to have a list.
         ' set the recordset up and pass the empty recordset onto the Build Recordset routine
        Set recTemp = New ADODB.Recordset
        BuildRecordset recTemp, recProviders
    End If
'R. Foster SCR #34505;   Developer: Robert Foster 04/24/2003 10:57 AM
'R. Foster Note: Function Changed:   CHCSII_NonProvider.dll\OpenNonProvider - OpenNonProvider; getProviders   CHCSII_Consult.dll\Form_Load      Impact of Change: Enable Saves to for NPOE in Provider GUI mode.
'Change for CHCS_GUI...
Else 'We're in CHCS_GUI mode so work appropriately...
    strFront = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
    strFront = strFront & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "

    If gobjColl.Count <> 0 Then ' Add the OSH items to the Select Statement
        strTemp = "WHERE NCID IN ("
        For iCount = 1 To gobjColl.Count
            If iCount = 1 Then
                strTemp = strTemp & gobjColl(iCount)
            Else
                strTemp = strTemp & ", " & gobjColl(iCount)
            End If
        Next iCount
        strTemp = strTemp & ")"
        If strTemp <> "WHERE NCID IN ()" Then
            strSQL = strFront & strTemp & " AND FACILITY_NCID = " & glFacilityNCID
            strSQL = strSQL & " AND SIGNATURE_CLASS IN(2,3,4)"
        End If
        Set recTemp = objDas.OpenRecordset(strSQL)
        BuildRecordset recTemp, recProviders
    Else 'This is a non-provider in CHCS_GUI mode and they have never chosen any providers to have a list.
         ' set the recordset up and pass the empty recordset onto the Build Recordset routine
        Set recTemp = New ADODB.Recordset
        BuildRecordset recTemp, recProviders
    End If

End If


End Function
Private Function ReOrderCollection(lngItemData As Long)
'This is where we reorder the collection so that we can move the person that was selected to the top
'This avoids a keying problem that will arise as you move things around.
Dim colTemp As New Collection
'Dim strName As String
Dim iCount As Integer
Dim strNCID As String
'strName = "Non_Providers_OE"
strNCID = gobjColl(lngItemData)
'Add the selected provider to the Temporary Collection first
colTemp.Add gobjColl(lngItemData), strName & " - " & colTemp.Count + 1
'Now remove the selected provider from the global collection
gobjColl.Remove lngItemData
'Now add the rest of the providers to colTemp
For iCount = 1 To gobjColl.Count
    colTemp.Add gobjColl(iCount), strName & " - " & colTemp.Count + 1
Next iCount
'Now remove the rest from the global
For iCount = 1 To gobjColl.Count
    gobjColl.Remove 1
Next iCount
'now add them all back in
For iCount = 1 To colTemp.Count
    gobjColl.Add colTemp(iCount), strName & " - " & gobjColl.Count + 1
    lngCollAdds = gobjColl.Count
Next iCount

End Function
Public Function AddProviderToListByNCID(recProviders As ADODB.Recordset, lngNCID As Long) As Boolean
Dim objDas As ICHCSII_DAS
Dim strCheck As String
Dim recTemp As ADODB.Recordset
'Dim strName As String
'strName = "Non_Providers_OE"
Set objDas = SetCHCSII_DAS
Select Case gintSig_Class
    Case 2, 3, 4 'They are a signable provider...
    
    Case Else

        'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:57 AM
        strCheck = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
        'Robert Foster End: SCR #23369;
        strCheck = strCheck & "WHERE CLINICIAN_NCID = " & lngNCID
        'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:20 PM
        'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
        strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
        'RFoster End: SCR #
        Set recTemp = objDas.OpenRecordset(strCheck)
        If recTemp.RecordCount = 0 Then
            recTemp.Close
            'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:58 AM
            strCheck = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
            'Robert Foster End: SCR #23369;
            strCheck = strCheck & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
            strCheck = strCheck & "WHERE NCID=" & lngNCID
            'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:20 PM
            'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
            strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
            'RFoster End: SCR #
            Set recTemp = objDas.OpenRecordset(strCheck)
            If Not recTemp.EOF Then
                recTemp.MoveFirst
                If gobjColl.Count = 0 Then ' add it or add it to the first position...
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1
                    lngCollAdds = gobjColl.Count
                Else
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1, 1
                    lngCollAdds = gobjColl.Count
                End If
                recTemp.Close
                'save the changes to the OSH based on the possible changes
                Call ResetOSH
                'if they changed something refetch the providers
                Call getProviders(recProviders)
                AddProviderToListByNCID = True
            Else
                AddProviderToListByNCID = False
                recTemp.Close
                Set recTemp = Nothing
            End If
        Else
        
            If Not recTemp.EOF Then
                recTemp.MoveFirst
                If gobjColl.Count = 0 Then ' add it or add it to the first position...
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1
                    lngCollAdds = gobjColl.Count
                Else
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1, 1
                    lngCollAdds = gobjColl.Count
                End If
                recTemp.Close
                'save the changes to the OSH based on the possible changes
                Call ResetOSH
                'if they changed something refetch the providers
                Call getProviders(recProviders)
                AddProviderToListByNCID = True
            End If
        
        End If
End Select
Set objDas = Nothing
Set recTemp = Nothing
End Function

Public Function AddAdditionalProviderToListByNCID(recProviders As ADODB.Recordset, lngNCID As Long) As Boolean
'R. Foster SCR #  Developer: Robert Foster 02/21/2003 06:37 PM
'R. Foster Note: Function Changed:   CLSNonProvider: AddAdditionalProviderToListByNCID; AdditionalProviderSelected  ProviderList.frm: Declarations; Display      Impact of Change: Added function to Add pre-chosen provider to list in the event that the MRU list does not contain this provider.  Added code to AdditionalProviderSelected to make sure that the collection has provider.  Added a check to avoid running the Provider combo click event during Display
Dim objDas As ICHCSII_DAS
Dim strCheck As String
Dim recTemp As ADODB.Recordset
Dim vItem As Variant
Dim bFound As Boolean
Dim lngItemData As Long
'strName = "Non_Providers_OE"
Set objDas = SetCHCSII_DAS
'get the additional providers list.
'R. Foster SCR #32633;   Developer: Robert Foster 02/24/2003 02:44 PM
'R. Foster Note: Function Changed:  Providerlist.frm cboProv_Click; Display  clsNonProvider AddAditionalProviderToListByNCID    Impact of Change: Modified selection criteria to drop down; modified combo style to limit list to items in combo. aaded call to reset osh added call to geet additional providers to ensure list completion.
getAdditionalProvidersList recProviders




strCheck = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
strCheck = strCheck & "WHERE CLINICIAN_NCID = " & lngNCID
strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
Set recTemp = objDas.OpenRecordset(strCheck)
If recTemp.RecordCount = 0 Then
    recTemp.Close
    strCheck = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
    strCheck = strCheck & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
    strCheck = strCheck & "WHERE NCID=" & lngNCID
    strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
    Set recTemp = objDas.OpenRecordset(strCheck)
    If Not recTemp.EOF Then
        recTemp.MoveFirst
        If gobjColl.Count = 0 Then
                gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1
                lngCollAdds = gobjColl.Count
                Call ResetOSH
                Call getAdditionalProvidersList(recProviders)
                
        Else
            For Each vItem In gobjColl
                If CStr(vItem) = recTemp.Fields("CLINICIAN_NCID").Value Then
                    bFound = True
                    Exit For
                End If
            Next vItem
            If Not bFound Then
                gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & lngCollAdds + 1, 1
                lngCollAdds = lngCollAdds + 1
                AddAdditionalProviderToListByNCID = True
                Call ResetOSH
                Call getAdditionalProvidersList(recProviders)
            Else ' they searched for one already in their list...
                recProviders.MoveFirst
                recProviders.Find "NCID=" & recTemp.Fields("CLINICIAN_NCID").Value
                lngItemData = recProviders.Fields("FLD_ORDER").Value
                Call AdditionalProviderSelected(recProviders, lngItemData)
                AddAdditionalProviderToListByNCID = True
                Exit Function
            End If
        End If
    Else
        AddAdditionalProviderToListByNCID = False
        recTemp.Close
        Set recTemp = Nothing
    End If
Else

    If Not recTemp.EOF Then
        recTemp.MoveFirst
        If gobjColl.Count = 0 Then
                gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1
                lngCollAdds = gobjColl.Count
                'R. Foster SCR #32633;   Developer: Robert Foster 02/24/2003 02:45 PM
                'R. Foster Note: Function Changed:  Providerlist.frm cboProv_Click; Display  clsNonProvider AddAditionalProviderToListByNCID    Impact of Change: Modified selection criteria to drop down; modified combo style to limit list to items in combo. aaded call to reset osh added call to geet additional providers to ensure list completion.
                Call ResetOSH
                Call getAdditionalProvidersList(recProviders)
               
        Else
            For Each vItem In gobjColl
                If CStr(vItem) = recTemp.Fields("CLINICIAN_NCID").Value Then
                    bFound = True
                    Exit For
                End If
            Next vItem
            If Not bFound Then
                gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & lngCollAdds + 1, 1
                lngCollAdds = lngCollAdds + 1
                AddAdditionalProviderToListByNCID = True
                'R. Foster SCR #32633;   Developer: Robert Foster 02/24/2003 02:45 PM
                'R. Foster Note: Function Changed:  Providerlist.frm cboProv_Click; Display  clsNonProvider AddAditionalProviderToListByNCID    Impact of Change: Modified selection criteria to drop down; modified combo style to limit list to items in combo. aaded call to reset osh added call to geet additional providers to ensure list completion.
                Call ResetOSH
                Call getAdditionalProvidersList(recProviders)
            Else ' they searched for one already in their list...
                recProviders.MoveFirst
                recProviders.Find "NCID=" & recTemp.Fields("CLINICIAN_NCID").Value
                lngItemData = recProviders.Fields("FLD_ORDER").Value
                Call AdditionalProviderSelected(recProviders, lngItemData)
                AddAdditionalProviderToListByNCID = True
                Exit Function
            End If
        End If
        AddAdditionalProviderToListByNCID = True
    End If

End If
Set objDas = Nothing
Set recTemp = Nothing
End Function

Public Function AddProviderToList(recProviders As ADODB.Recordset, lngIEN As Long) As Boolean
Dim objDas As ICHCSII_DAS
Dim strCheck As String
Dim recTemp As ADODB.Recordset
'Dim strName As String
'strName = "Non_Providers_OE"
Set objDas = SetCHCSII_DAS
Select Case gintSig_Class
    Case 2, 3, 4 'They are a signable provider...
    
    Case Else
    
        'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:58 AM
        strCheck = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
        'Robert Foster End: SCR #23369;
        strCheck = strCheck & "WHERE PROVIDER_IEN = " & lngIEN
        'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:20 PM
        'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
        strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
        'RFoster End: SCR #
        Set recTemp = objDas.OpenRecordset(strCheck)
        If recTemp.RecordCount = 0 Then
            recTemp.Close
            'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:58 AM
            strCheck = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
            'Robert Foster End: SCR #23369;
            strCheck = strCheck & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
            strCheck = strCheck & "WHERE IEN=" & lngIEN
            'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:21 PM
            'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
            strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
            'RFoster End: SCR #
            Set recTemp = objDas.OpenRecordset(strCheck)
            If Not recTemp.EOF Then
                recTemp.MoveFirst
                If gobjColl.Count = 0 Then ' add it or add it to the first position...
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1
                    lngCollAdds = gobjColl.Count
                Else
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1, 1
                    lngCollAdds = gobjColl.Count
                End If
                recTemp.Close
                'save the changes to the OSH based on the possible changes
                Call ResetOSH
                'if they changed something refetch the providers
                Call getProviders(recProviders)
                AddProviderToList = True
            Else
                AddProviderToList = False
                recTemp.Close
                Set recTemp = Nothing
            End If
        Else
        
            If Not recTemp.EOF Then
                recTemp.MoveFirst
                If gobjColl.Count = 0 Then ' add it or add it to the first position...
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1
                    lngCollAdds = gobjColl.Count
                Else
                    gobjColl.Add recTemp.Fields("CLINICIAN_NCID").Value, strName & " - " & gobjColl.Count + 1, 1
                    lngCollAdds = gobjColl.Count
                End If
                recTemp.Close
                'save the changes to the OSH based on the possible changes
                Call ResetOSH
                'if they changed something refetch the providers
                Call getProviders(recProviders)
            End If
        
        End If
End Select
    Set objDas = Nothing
    Set recTemp = Nothing
    End Function

Public Function ProviderSelected(recProviders As ADODB.Recordset, lngItemData As Long)
'Dim strName As String
Dim iCount As Integer
Dim strNCID As String
'strName = "Non_Providers_OE"
strNCID = gobjColl(lngItemData)
'Put the selection to the top of the collection
Call ReOrderCollection(lngItemData)
Call ResetOSH

recProviders.MoveFirst
recProviders.Find "NCID=" & strNCID
'Now set up the recordset...
Call getProviders(recProviders)

End Function

Public Function AdditionalProviderSelected(recProviders As ADODB.Recordset, lngItemData As Long)
'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:11 PM
'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode
'Dim strName As String
Dim iCount As Integer
Dim strNCID As String
'strName = "Non_Providers_OE"
'R. Foster SCR #  Developer: Robert Foster 02/21/2003 06:37 PM
'R. Foster Note: Function Changed:   CLSNonProvider: AddAdditionalProviderToListByNCID; AdditionalProviderSelected  ProviderList.frm: Declarations; Display      Impact of Change: Added function to Add pre-chosen provider to list in the event that the MRU list does not contain this provider.  Added code to AdditionalProviderSelected to make sure that the collection has provider.  Added a check to avoid running the Provider combo click event during Display
If gobjColl.Count > 0 And lngItemData <> 0 Then
    strNCID = gobjColl(lngItemData)
    'Put the selection to the top of the collection
    Call ReOrderCollection(lngItemData)
    Call ResetOSH
    
    recProviders.MoveFirst
    recProviders.Find "NCID=" & strNCID
    'Now set up the recordset...
    Call getAdditionalProvidersList(recProviders)
End If
End Function

Public Function OpenNonProvider(robjLogon As Object, robjComm As Object, robjUser As Object) As Object
'Robert Foster Begin:SCR # 260 (Pharmacy) Developer: Robert Foster 06/06/2002 01:52 PM
'Robert Foster Note: Function Changed:      Impact of Change:
'This is where we really intitialize the object
    Dim objOshMan As New OshMan
    Dim objDas As ICHCSII_DAS
    Dim recTemp As ADODB.Recordset
    Dim strCheck As String
    Dim intTemp As Integer
    'RFoster Begin:SCR #27031;   Developer: Robert Foster 08/26/2002 03:14 PM
    'RFoster Note: Function Changed: OpenNonProvider      Impact of Change: This changes the Provider  selection from Last used provider (as per original direction) to the encounter's primary provider (as per latest stated preference).  When the non-provider launches an Order Entry Module (and Launches NPOE), NPOE dll will get the current encounter and set the default choice to that provider.
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim objCurrentEncounter As CHCSIIEncounterOps.Encounter
    Dim lngtItem As Long
    Dim bFound As Boolean
    
    'RFoster End: SCR #27031;
    'cwobjConfig is used to retrieve and save values from the OSH.
    Set objDas = SetCHCSII_DAS
    Set cwobjUser = robjUser
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogon
'RFoster Begin:SCR #27031;   Developer: Robert Foster 08/26/2002 03:15 PM
'RFoster Note: Function Changed: OpenNonProvider      Impact of Change: This changes the Provider  selection from Last used provider (as per original direction) to the encounter's primary provider (as per latest stated preference).  When the non-provider launches an Order Entry Module (and Launches NPOE), NPOE dll will get the current encounter and set the default choice to that provider.
Set objEncParent = cwobjCWComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
If Not objEncParent Is Nothing Then
    Set objCurrentEncounter = objEncParent.CurrentEncounter
    lintEncounterProvider = objCurrentEncounter.PrimaryProvider.NCID
End If
'We're done with them; lose 'em.
Set objCurrentEncounter = Nothing
Set objEncParent = Nothing
    'RFoster End: SCR #27031;
    glNursingDivisionNCID = Val(robjLogon.NursingDivisionNCID)
    glFacilityNCID = Val(robjLogon.FacilityNCID)
    glUserNCID = Val(cwobjLogonEnviron.UserNCID)
    'Set the global signature class for the provider
    'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:59 AM
    strCheck = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
    'Robert Foster End: SCR #23369;
    strCheck = strCheck & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
    strCheck = strCheck & "WHERE NCID=" & glUserNCID
    'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:21 PM
    'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
    strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
'RFoster End: SCR #
'    strCheck = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
'    strCheck = strCheck & "WHERE CLINICIAN_NCID = " & glUserNCID
    'for some reason, when you look up witht he CLINIC_NCID you might not find it
    ' strCheck = strCheck & " AND FACILITY_NCID = " & glfacilityNCID
    Set recTemp = objDas.OpenRecordset(strCheck)
    
    If recTemp.RecordCount = 0 Then
        recTemp.Close
        'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:59 AM
        'R. Foster SCR #31265; 31398;   Developer: Robert Foster 01/13/2003 04:58 PM
        'R. Foster Note: Function Changed: OpenNonProvider      Impact of Change: re-write SQL Statement to properly retrieve Provider from the Clinic_Provider table when Provider is missing from the Provider Table.
        strCheck = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
        'Robert Foster End: SCR #23369;
        strCheck = strCheck & "WHERE CLINICIAN_NCID = " & glUserNCID
        'for some reason, when you look up witht he CLINIC_NCID you might not find it
        ' strCheck = strCheck & " AND FACILITY_NCID = " & glfacilityNCID
        'Robert Foster Begin:SCR #23369;   Developer: RFoster 05/16/2002 09:59 AM
        'strCheck = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
        'Robert Foster End: SCR #23369;
        'strCheck = strCheck & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
        'strCheck = strCheck & "WHERE CLINICIAN_NCID=" & glUserNCID
        'RFoster Begin:SCR #  Developer: Robert Foster 08/23/2002 02:21 PM
        'RFoster Note: Function Changed: (clnsrch) Display & (clsNonProvider) OpenNonProvider, getProviders,  AddProviderToListByNCID, AddProviderToList,       Impact of Change: Liumits Non-Provider Order Entry Search to the users MTF.
        strCheck = strCheck & " AND FACILITY_NCID = " & glFacilityNCID
        'RFoster End: SCR #
        Set recTemp = objDas.OpenRecordset(strCheck)
    End If

    
    If recTemp.RecordCount = 1 Then  'If there is more than one then "...Houston we have a problem..."
        If Not IsNull(recTemp.Fields("SIGNATURE_CLASS").Value) And Trim(recTemp.Fields("SIGNATURE_CLASS").Value) <> "" Then
            gintSig_Class = recTemp.Fields("SIGNATURE_CLASS").Value
        End If
    Else
        If recTemp.RecordCount > 1 Then 'We have a problem so go and get the greatest value
            recTemp.MoveFirst
            Do Until recTemp.EOF
            'Obviously, if we have more than one, then we need to find the CLASS_SIGNATURE
            'and they exist with more than one SIGNATURE_CLASS in the Dev Environment
            'meaning that it's possible, but I hope not.
                If Not IsNull(recTemp.Fields("SIGNATURE_CLASS").Value) And Trim(recTemp.Fields("SIGNATURE_CLASS").Value) <> "" Then
                    If recTemp.Fields("SIGNATURE_CLASS").Value > intTemp Then
                        intTemp = recTemp.Fields("SIGNATURE_CLASS").Value
                        gintSig_Class = intTemp
                    End If
                End If
                recTemp.MoveNext
            Loop
        Else
            gintSig_Class = 0
        End If
    End If

    'Look in the OSH for the IEN of the users that were chosen...
    'Added If check for ADM Demo by Phil Crowder 12/5/02
    If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR 37894
      Set gobjColl = objOshMan.GetOshItems(strName, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID))
    'R. Foster SCR #34505;   Developer: Robert Foster 04/24/2003 10:57 AM
    'R. Foster Note: Function Changed:   CHCSII_NonProvider.dll\OpenNonProvider - OpenNonProvider; getProviders   CHCSII_Consult.dll\Form_Load      Impact of Change: Enable Saves to for NPOE in Provider GUI mode.
    Else
      objOshMan.ConfigObject = mobjConfig
      Set gobjColl = objOshMan.GetOshItems(strName, strDataType & "-" & CStr(glUserNCID), CStr(glUserNCID))
    End If

    If gobjColl Is Nothing Then '...Nothing was found in the OSH so no items in the list
        Set gobjColl = New Collection
    End If
    'Add the EncounterProvider to the object if necessary...
    
       'Robert Foster Begin:SCR # 260 (Pharmacy) Developer: Robert Foster 06/06/2002 01:54 PM
   'Robert Foster Note: Function Changed: OpenNonProvider      Impact of Change: Ability to utilize SignableProvider boolean without using GetProviders (which requires a recordset argument)
   'SCR#: 260 Pharmacy
    Select Case gintSig_Class
    'At present the only signature clases that are defined as Signable Providers
        Case 2, 3, 4
            mboolSignableProvider = True
    'All other are defined as non-providers; all that should show here are classes 0 & 1
    'Class 99 are outside providers and do not have access.  They have been filtered out as a result.
        Case Else
            mboolSignableProvider = False
    End Select
   'Robert Foster End: SCR #

    
    'RFoster Begin:SCR #27031;   Developer: Robert Foster 08/26/2002 03:15 PM
    'RFoster Note: Function Changed: OpenNonProvider      Impact of Change: This changes the Provider  selection from Last used provider (as per original direction) to the encounter's primary provider (as per latest stated preference).  When the non-provider launches an Order Entry Module (and Launches NPOE), NPOE dll will get the current encounter and set the default choice to that provider.
    bFound = False
    If mboolSignableProvider = False Then
        For lngtItem = 1 To gobjColl.Count
            If gobjColl(lngtItem) = lintEncounterProvider Then
                bFound = True
                Exit For
            End If
        Next lngtItem
        If Not bFound Then
            If gobjColl.Count <> 0 Then
                gobjColl.Add lintEncounterProvider, strName & " - " & gobjColl.Count + 1, 1
                lngCollAdds = gobjColl.Count
            Else
                gobjColl.Add lintEncounterProvider, strName & " - " & gobjColl.Count + 1
                lngCollAdds = gobjColl.Count
            End If
        Else
            Call ReOrderCollection(lngtItem)
        End If
    End If
    'RFoster End: SCR #27031;
    
    recTemp.Close
    Set recTemp = Nothing
    Set objOshMan = Nothing
    Set objDas = Nothing
    
    Set OpenNonProvider = Me

End Function
Private Function BuildRecordset(recTemp As ADODB.Recordset, recProviders As ADODB.Recordset) As Boolean
Dim iCount As Integer
'RFoster Begin:SCR #27142;   Developer: Robert Foster 09/25/2002 12:31 PM
'RFoster Note: Function Changed: BuildRecordset      Impact of Change: In certain instances the recordset is getting corrupted; make sure it is set to New to compensate for this
Set recProviders = New ADODB.Recordset
recProviders.CursorLocation = adUseClient
recProviders.CursorType = adOpenStatic
'RFoster End: SCR #27142;
'get rid of the fields in the recordset...
 For iCount = 0 To recProviders.Fields.Count - 1
     recProviders.Fields.Delete 0
 Next
'...and re-append them
 recProviders.Fields.Append fldORDER, adInteger
 recProviders.Fields.Append fldIEN, adDouble
 recProviders.Fields.Append fldNCID, adDouble
 recProviders.Fields.Append fldProviderName, adLongVarChar, 30
 recProviders.Fields.Append fldSignatureClass, adInteger
 recProviders.Open
 
'Now add the values to the recordset
If recTemp.State <> adStateClosed Then
'    If Not recTemp.EOF And Not recTemp.BOF Then
    If Not recTemp.RecordCount = 0 Then
        For iCount = 1 To gobjColl.Count
            
            recTemp.MoveFirst
            recTemp.Find "CLINICIAN_NCID=" & gobjColl(iCount)
            If Not recTemp.EOF Then
                recProviders.AddNew
                recProviders.Fields("FLD_ORDER").Value = iCount
                recProviders.Fields("IEN").Value = recTemp.Fields("PROVIDER_IEN").Value
                recProviders.Fields("NCID").Value = recTemp.Fields("CLINICIAN_NCID").Value
                recProviders.Fields("NAME").Value = recTemp.Fields("PROVIDER_NAME").Value & ""
                recProviders.Fields("Signature_Class").Value = recTemp.Fields("Signature_Class").Value
                recProviders.Update
            End If
            
        Next iCount
        If recProviders.RecordCount > 1 Then
            recProviders.Sort = "FLD_ORDER"
        End If
    End If
End If
If recProviders.RecordCount <> 0 Then
    recProviders.MoveFirst
End If
'Do Until recProviders.EOF
'    Debug.Print recProviders.Fields(0).Value & "~" & recProviders.Fields(1).Value & "~" & recProviders.Fields(2).Value & "~" & recProviders.Fields(3).Value
'    recProviders.MoveNext
'Loop
End Function
Private Sub Class_Initialize()
'  Debug.Print "CHCSII_NonProvider Class_Initialize"
  
#If coiDebug Then
     gCOI_DL.debugStart
     DR.TimerTime = 0
     gbDebugStarted = True
'     gCOI_DL.DebugLog "COI31", "ALERT", "", "Begin", DR
#End If

    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    'Set cwobjNavigation = New Navigation

    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    
    Set m_objCHCSII_DAS = SetCHCSII_DAS
    Set gobjColl = New Collection
    mboolFirstTime = True
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
    


End Sub

Private Sub Class_Terminate()
'reset the OSH an capture any of the modifications that have taken place
'Call ResetOSH
'glNursingDivisionNCID = 0
'glFacilityNCID = 0
'glUserNCID = 0
'gintSig_Class = 0


  Set m_objCHCSII_DAS = Nothing
  Set gobjColl = Nothing
'  Debug.Print "CHCSII_NonProvider Class_Terminate"
End Sub
