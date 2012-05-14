VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncAccidentData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NCID_NoAccident = 157059    '3M CID = No Accident- Not Accident Related
Private Const NCID_Accident = 157060      '3M CID = Accident- Accident Related

Private Const MODULE_NAME = "EncAccidentData"
Private Const EMPTY_LOCATION_IEN = 0
Private Const NCID_DELIMETER = ","
Private Const MSG_INVALID_CAUSE_CODE = "An accident cause code must be specified."
Private Const MSG_INVALID_LOCATION = "A location must be specified when the cause code 'AA - Auto Accident' is selected."
Private Const MSG_INVALID_TOO_MANY_CAUSE_CODES = "No more than three (3) cause codes may be selected."

'Accident Cause Code NCIDs Enumeration
Public Enum AccidentCauseCodes
   accNotSpecified = -1
   accAutoAccident = 555
   accAnotherPartyResponsible = 1276
   accEmploymment = 558
   accOtherAccident = 144
End Enum

'Declare Object State Variables
Private mbDirty As Boolean
Private mbInDB As Boolean                 'Indicates if the record exists in the database

'Declare Data Variables
Private mdAccidentDate As Date
Private mbCausedByAccident As Boolean
Private msCauseCodeNCID As String
Private msLocationIEN As String           'Number(20) - too big for long
Private msLocationCode As String          'Location Abbreviation
Private msApptID As String                'Number(20) - too big for long
Private msClinicNCID As String            'Encounter Clinic NCID

''''''''''''''''''''''''''''''''''''''''''''''
' Property Values
''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get IsDirty() As Boolean
   IsDirty = mbDirty
End Property

Public Property Get AccidentDate() As Date
   AccidentDate = mdAccidentDate
End Property

Public Property Let AccidentDate(ByVal dNewValue As Date)
   If mdAccidentDate <> dNewValue Then
      mdAccidentDate = dNewValue
      mbDirty = True
   End If
End Property

Public Property Get CausedByAccident() As Boolean
   CausedByAccident = mbCausedByAccident
End Property

Public Property Let CausedByAccident(ByVal bNewValue As Boolean)
   If mbCausedByAccident <> bNewValue Then
      mbCausedByAccident = bNewValue
      If Not mbCausedByAccident Then Clear
      mbDirty = True
   End If
End Property

Public Property Get CauseCodeNCID() As String
   CauseCodeNCID = msCauseCodeNCID
End Property

Public Property Let CauseCodeNCID(ByVal sNewValue As String)

On Error GoTo ErrorHandler

Dim vCauseCodes() As String
Dim lIdx As Long
   
   If msCauseCodeNCID <> sNewValue Then
      vCauseCodes = Split(sNewValue, NCID_DELIMETER)
      For lIdx = 0 To UBound(vCauseCodes)
         Select Case vCauseCodes(lIdx)
         Case accNotSpecified, accAutoAccident, accEmploymment, accAnotherPartyResponsible, accOtherAccident
            'Value is ok
         Case Else
            Err.Raise 911, , "Invalid Cause Code NCID."
         End Select
      Next
      
      msCauseCodeNCID = sNewValue
      mbDirty = True
   End If
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CauseCodeNCID"
   
End Property

Public Property Get ClinicNCID() As String
   ClinicNCID = msClinicNCID
End Property

Public Property Get LocationCode() As String
   LocationCode = msLocationCode
End Property

Public Property Let LocationCode(ByVal sNewValue As String)
   'This property is not persisted so don't set the dirty flag of object when changed
   msLocationCode = sNewValue
End Property

Public Property Get LocationIEN() As String
   LocationIEN = msLocationIEN
End Property

Public Property Let LocationIEN(ByVal sNewValue As String)
   If Not IsNumeric(sNewValue) Then
      Err.Raise 911, , "Location IEN must be numeric."
   ElseIf Val(sNewValue) < EMPTY_LOCATION_IEN Then
      Err.Raise 911, , "Location IEN must be a positive number."
   ElseIf msLocationIEN <> sNewValue Then
      msLocationIEN = sNewValue
      mbDirty = True
   End If
End Property

Private Property Get LocationName(ByVal sFacilityNCID As String) As String

On Error GoTo ErrorHandler

Dim oListMgr As actxDODCentricCodes.ListManagerOps
Dim oLocation As actxDODCentricCodes.Location

   'Initialize Return Value
   LocationName = vbNullString
   
   If LocationIEN <> EMPTY_LOCATION_IEN Then
      Set oListMgr = New actxDODCentricCodes.ListManagerOps
      For Each oLocation In oListMgr.GetFacilityLocations(CLng(sFacilityNCID))
         If oLocation.IEN = LocationIEN Then
            LocationName = oLocation.Name
            Exit For
         End If
      Next
   End If
   
ErrorHandler:
   'Clean Up Objects
   Set oListMgr = Nothing
   Set oLocation = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "LocationName"
   
End Property

Public Property Get Completed() As Boolean
   'If cause code is set then accident info has been completed
   Completed = CBool(msCauseCodeNCID <> accNotSpecified)
End Property

'Validates the accident data - returns a warning message if the data is not valid
Public Function Validate(ByRef sWarning As String) As Boolean

On Error GoTo ErrorHandler

   If mbCausedByAccident And msCauseCodeNCID = accNotSpecified Then
      sWarning = MSG_INVALID_CAUSE_CODE
   ElseIf mbCausedByAccident And IsCausedBy(accAutoAccident) And (LocationIEN = EMPTY_LOCATION_IEN Or LocationIEN = vbNullString) Then
      sWarning = MSG_INVALID_LOCATION
   ElseIf UBound(Split(msCauseCodeNCID, NCID_DELIMETER)) > 2 Then
      sWarning = MSG_INVALID_TOO_MANY_CAUSE_CODES
   Else
      Validate = True
      sWarning = vbNullString
   End If

ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "Validate"

End Function

'Checks to see if the accident was cause (at least in part) by the specified cause code NCID
Public Function IsCausedBy(ByVal eCauseCodeNCID As AccidentCauseCodes) As Boolean
   
   'Check cause code list to see if it contains the value passed in
   IsCausedBy = CBool(InStr(1, msCauseCodeNCID, CStr(eCauseCodeNCID)) > 0)
   
End Function

'Provide the select field list for encounter load
Friend Function GetSQLFields() As String
   GetSQLFields = " EA.ACCIDENT_INDICATOR, EA.ACCIDENT_DATE, EA.LOCATION_IEN, EA.ACCIDENT_CAUSE_CODE_NCID "
End Function

Friend Function GetSQLValues() As String
   If CausedByAccident Then
      'Set accident field values
      GetSQLValues = " 1, " & _
                     gobjCHCSIIConn.SQLDate(AccidentDate) & ", " & _
                     IIf(LocationIEN = EMPTY_LOCATION_IEN, "NULL", LocationIEN) & ", " & _
                     SQLQuote(CauseCodeNCID) & ", "
   Else
      'Clear accident field values
      GetSQLValues = "VALUES (0, NULL, NULL, NULL )"
   End If
End Function

Friend Function GetSQLUpdate() As String
   If CausedByAccident Then
      GetSQLUpdate = ", ACCIDENT_INDICATOR = 1" & _
                     ", ACCIDENT_DATE = " & gobjCHCSIIConn.SQLDate(AccidentDate) & _
                     ", ACCIDENT_LOCATION_IEN = " & IIf(LocationIEN = EMPTY_LOCATION_IEN, "NULL", LocationIEN) & _
                     ", ACCIDENT_CAUSE_CODE_NCID = " & SQLQuote(CauseCodeNCID) & ", "
   Else
      GetSQLUpdate = ", ACCIDENT_INDICATOR = 0" & _
                     ", ACCIDENT_DATE = NULL" & _
                     ", ACCIDENT_LOCATION_IEN = NULL" & _
                     ", ACCIDENT_CAUSE_CODE_NCID = NULL"
   End If
End Function

Private Sub Class_Initialize()
   Clear
End Sub

'Initialize data object from SQL object
Friend Sub InitFromSQL(ByVal sClinicNCID As String, Optional ByRef oSQLEx As ICHCSII_SQL.ISqlOps, _
      Optional ByVal sApptId As String = vbNullString)

On Error GoTo ErrorHandler

Dim sSql As String
   
   If oSQLEx Is Nothing And Len(sApptId) > 0 Then
      'Load directly from table
      Set oSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
      oSQLEx.Execute "SELECT EA.APPT_ID as APPTID, " & GetSQLFields & _
            " FROM Encounter_Accident EA " & _
            " WHERE EA.APPT_ID = " & sApptId
   End If
      
   If Not oSQLEx.EOF Then
      msApptID = oSQLEx("APPTID") & vbNullString
      CausedByAccident = CBool(Val(oSQLEx("ACCIDENT_INDICATOR") & vbNullString) = 1)
      AccidentDate = IIf(oSQLEx("ACCIDENT_DATE") & vbNullString = vbNullString, Format(Now, "MM/DD/YYYY"), oSQLEx("ACCIDENT_DATE"))
      LocationIEN = IIf(oSQLEx("LOCATION_IEN") & vbNullString = vbNullString, EMPTY_LOCATION_IEN, oSQLEx("LOCATION_IEN"))
      CauseCodeNCID = IIf(oSQLEx("ACCIDENT_CAUSE_CODE_NCID") & vbNullString = vbNullString, accNotSpecified, oSQLEx("ACCIDENT_CAUSE_CODE_NCID"))
      mbInDB = CausedByAccident
   Else
      Clear
      msApptID = sApptId
      mbInDB = False
   End If
   
   msClinicNCID = sClinicNCID
   mbDirty = False
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "InitFromSQL"
   
End Sub

'Persist the data to the database
'must pass in NewApptId for new appointments or facility and enc number
Public Function GemsSave(Optional ByVal sNewApptId As String, _
      Optional ByVal sFacilityNCID As String, Optional ByVal sEncNumber As String) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim oDas As ICHCSII_DAS

   If msApptID = vbNullString Or Not mbInDB Then
      'This is a new record since no apptID was loaded from the database
      If CausedByAccident Then
         'Insert new record if caused by accident
         sSql = "INSERT INTO Encounter_Accident " & _
                  "( APPT_ID, ACCIDENT_INDICATOR, " & _
                  " ACCIDENT_DATE, LOCATION_IEN, ACCIDENT_CAUSE_CODE_NCID )" & _
               " VALUES ( " & IIf(sNewApptId = vbNullString, msApptID, sNewApptId) & ", " & " 1, " & _
                     gobjCHCSIIConn.FormatDateForDBInsert(AccidentDate) & ", " & _
                     IIf(LocationIEN = EMPTY_LOCATION_IEN, "NULL", LocationIEN) & ", " & _
                     SQLQuote(CauseCodeNCID) & ")"
      End If
   ElseIf CausedByAccident Then
      'Update Existing Record
      sSql = "UPDATE Encounter_Accident SET " & _
         " EncounterNumber = " & sEncNumber & ", " & _
         " FacilityNCID = " & sFacilityNCID & ", " & _
         " Accident_Indicator = 1" & ", " & _
         " Accident_Date = " & gobjCHCSIIConn.FormatDateForDBInsert(AccidentDate) & ", " & _
         " Location_Ien = " & IIf(LocationIEN = EMPTY_LOCATION_IEN, "null", LocationIEN) & ", " & _
         " Accident_Cause_Code_NCID = " & SQLQuote(CauseCodeNCID) & _
         " WHERE Appt_Id = " & msApptID
   Else
      'Not caused by accident so remove previous data
      sSql = "DELETE FROM Encounter_Accident WHERE Appt_Id = " & msApptID
   End If
   
   If sSql <> vbNullString Then
      Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto)
      oDas.ExecuteSQL sSql
      mbInDB = CBool(Left(sSql, 6) <> "DELETE")
   End If
   
   If Not CausedByAccident Then Clear
   
   mbDirty = False
   
   GemsSave = True
   
ErrorHandler:
   Set oDas = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GemsSave"
   
End Function


Private Function GetWAMCauseCode(ByVal sCauseCodeNCID As String) As String
   Select Case sCauseCodeNCID
   Case accAutoAccident
      GetWAMCauseCode = "AA"
   Case accAnotherPartyResponsible
      GetWAMCauseCode = "AP"
   Case accEmploymment
      GetWAMCauseCode = "EM"
   Case accOtherAccident
      GetWAMCauseCode = "OA"
   Case Else
      GetWAMCauseCode = vbNullString
   End Select
End Function

Private Function GetWAMCauseCodeDesc(ByVal sCauseCodeNCID As String) As String
Select Case sCauseCodeNCID
   Case accAutoAccident
      GetWAMCauseCodeDesc = "Auto Accident"
   Case accAnotherPartyResponsible
      GetWAMCauseCodeDesc = "Another Party Responsible"
   Case accEmploymment
      GetWAMCauseCodeDesc = "Employment"
   Case accOtherAccident
      GetWAMCauseCodeDesc = "Other Accident"
   Case Else
      GetWAMCauseCodeDesc = vbNullString
   End Select
End Function

Friend Function CopyToWam(ByRef objWamData As WAMData.EncAccident, ByVal sFacilityNCID As String)

On Error GoTo ErrorHandler

Dim oListMgrOps As actxDODCentricCodes.ListManagerOps
Dim oLocation As actxDODCentricCodes.Location
Dim vCauseCodes() As String
Dim lIdx As Long

   'Set date and clear cause codes collection
   objWamData.AccidentIndicator = Me.CausedByAccident
   objWamData.CauseCodes.Clear
   
   If Me.CausedByAccident Then
      objWamData.AccidentDate = Me.AccidentDate
      vCauseCodes = Split(Me.CauseCodeNCID, NCID_DELIMETER)
      For lIdx = 0 To UBound(vCauseCodes)
         If vCauseCodes(lIdx) <> accNotSpecified Then
            objWamData.CauseCodes.Add GetWAMCauseCode(vCauseCodes(lIdx)), GetWAMCauseCodeDesc(vCauseCodes(lIdx))
         End If
      Next
      
      objWamData.LocationIEN = Me.LocationIEN
      
      'SCR 50691 - Set the WAM location code - Lookup value if we don't have it
      If Len(Me.LocationCode) > 0 Then
         objWamData.LocationCode = Me.LocationCode
      ElseIf Me.LocationIEN <> EMPTY_LOCATION_IEN Then
         Set oListMgrOps = New actxDODCentricCodes.ListManagerOps
         For Each oLocation In oListMgrOps.GetFacilityLocations(sFacilityNCID)
            If oLocation.IEN = Me.LocationIEN Then
               Me.LocationCode = Trim(oLocation.Abbreviation)
               objWamData.LocationCode = Trim(oLocation.Abbreviation)
               Exit For
            End If
         Next
      Else
         objWamData.LocationCode = vbNullString
      End If
      
   Else
      'Clear out values
      objWamData.AccidentDate = NULL_DATE
      objWamData.LocationIEN = EMPTY_LOCATION_IEN
      objWamData.LocationCode = vbNullString
   End If
   
ErrorHandler:
   'Clean Up Objects
   Set oListMgrOps = Nothing
   Set oLocation = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CopyToWam"
   
End Function

Private Sub Clear()
   'Set or Reset Default Values
   mdAccidentDate = Format(Now, "MM/DD/YYYY")
   msLocationIEN = EMPTY_LOCATION_IEN
   msCauseCodeNCID = accNotSpecified
   mbCausedByAccident = False
   mbDirty = False
End Sub

Private Sub CDRAddCauseCode(ByRef oEnc As Occurrence, ByVal sCauseCode As String)
   oEnc.GenericData = "New GenericData"
   oEnc.GenericData.Type = sCauseCode
   oEnc.GenericData.ValueNcid = NCID_YES
End Sub

Private Sub CDRUpdateCauseCodes(ByRef oEnc As Occurrence)

   'DON'T use with keyword with genericdata object or it will produce strange comparison results
   oEnc.GenericData.GotoFirst
   
   Do While Not oEnc.GenericData.EOF
   
      If oEnc.GenericData.Type = accAutoAccident Or oEnc.GenericData.Type = accAnotherPartyResponsible _
         Or oEnc.GenericData.Type = accEmploymment Or oEnc.GenericData.Type = accOtherAccident Then
         
         If IsCausedBy(oEnc.GenericData.Type) Then
            oEnc.GenericData.ValueNcid = NCID_YES
         Else
            oEnc.GenericData.ValueNcid = NCID_NO
            oEnc.GenericData.Remove                    'This doesn't work but let's try anyway
         End If
         
      End If
      
      oEnc.GenericData.GoToNext
   
   Loop
         
End Sub

Public Function CDRLoad(ByVal sClinicNCID As String, ByVal sUnitNumber As String, ByVal sFacilityNCID As String, sEncNumber As String) As Boolean

On Error GoTo ErrorHandler

Dim oEnc As Occurrence
Dim lRet As Long
Dim sCauseCodes As String

   'Initialize the Occurrence Object
   Set oEnc = New Occurrence
   oEnc.EncounterNumber = sEncNumber
   oEnc.Facility = sFacilityNCID
   oEnc.UnitNumber = sUnitNumber
   
   lRet = oEnc.Read(False)
   If lRet > 0 Then
      
      Clear
            
      msClinicNCID = sClinicNCID
      Me.CausedByAccident = CBool(oEnc.AccidentCode <> "0")
      Me.AccidentDate = IIf(oEnc.AccidentTime = NULL_DATE, Format(Now, "MM/DD/YYYY"), oEnc.AccidentTime)
      Me.LocationIEN = IIf(oEnc.AccidentLocation = vbNullString, EMPTY_LOCATION_IEN, oEnc.AccidentLocation)
      
      If Not oEnc.GenericData Is Nothing Then
         Do While Not oEnc.GenericData.EOF
            Select Case oEnc.GenericData.Type
            Case accAutoAccident, accAnotherPartyResponsible, accEmploymment, accOtherAccident
               If oEnc.GenericData.ValueNcid = NCID_YES Then
                  sCauseCodes = sCauseCodes & "," & oEnc.GenericData.Type
               End If
            End Select
            oEnc.GenericData.GoToNext
         Loop
      End If
      
      If Len(sCauseCodes) > 0 Then CauseCodeNCID = Mid$(sCauseCodes, 2)
            
      mbDirty = False
      
      CDRLoad = True
      
   Else
      'Encounter not found
      Err.Raise 911, , "Failed to retrieve encounter occurence object for loading of accident data."
      
   End If
      
ErrorHandler:
   Set oEnc = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRLoad"
   
End Function

'Persist the data in CDR mode - saved in encounter_accident until encounter is created
'Must pass in NewApptId for new appointments or facility and enc number
Public Function CDRSave(ByVal sUnitNumber As String, ByVal sFacilityNCID As String, sEncNumber As String) As Boolean

On Error GoTo ErrorHandler

Dim oEnc As Occurrence
Dim lRet As Long
Dim vCauseCodes As Variant
Dim lIdx As Long
Dim bFound As Boolean

   'Initialize the Occurrence Object
   Set oEnc = New Occurrence
   oEnc.EncounterNumber = sEncNumber
   oEnc.Facility = sFacilityNCID
   oEnc.UnitNumber = sUnitNumber
   
   lRet = oEnc.Read(True)
   If lRet <> 0 Then
      'Update the encounter level accident fields
      'SCR 50480 - make sure we reset the cause code when it is not accident related
      If Me.CausedByAccident Then
         oEnc.AccidentCode = NCID_Accident
      Else
         oEnc.AccidentCode = NCID_NoAccident
      End If
      oEnc.AccidentTime = Format(Me.AccidentDate, "MM/DD/YYYY")
      oEnc.AccidentLocation = Me.LocationIEN
               
      'Add the cause codes
      If oEnc.GenericData Is Nothing Then
         'No cause codes or generic data items saved so add them
         vCauseCodes = Split(CauseCodeNCID, NCID_DELIMETER)
         For lIdx = 0 To UBound(vCauseCodes)
            CDRAddCauseCode oEnc, vCauseCodes(lIdx)
         Next lIdx
      Else
         'Clear any existing cause code values we have
         CDRUpdateCauseCodes oEnc
            
         'Add any missing cause code values
         vCauseCodes = Split(CauseCodeNCID, NCID_DELIMETER)
         For lIdx = 0 To UBound(vCauseCodes)
            oEnc.GenericData.GotoFirst          'DON'T use with keyword with genericdata object or it will produce strange comparison results
            bFound = False
            Do While Not oEnc.GenericData.EOF
               If oEnc.GenericData.Type = vCauseCodes(lIdx) _
                     And 0 = 0 Then
                  bFound = True
                  Exit Do
               End If
            oEnc.GenericData.GoToNext
            Loop
            If Not bFound Then CDRAddCauseCode oEnc, vCauseCodes(lIdx)
         Next lIdx
         
      End If
                  
      lRet = oEnc.WriteUpdate()
      
      If Not CausedByAccident Then Clear
      
      mbDirty = False
      
      CDRSave = CBool(lRet)
   
   Else
      'Could not find the encounter
      Err.Raise 911, , "Failed to retrieve encounter occurence object for loading of accident data."
   
   End If

ErrorHandler:
   Set oEnc = Nothing

   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "CDRSave"
   
End Function

Public Sub InitFromCopy(ByVal oSource As EncAccidentData)
   mbCausedByAccident = oSource.CausedByAccident
   mdAccidentDate = oSource.AccidentDate
   msCauseCodeNCID = oSource.CauseCodeNCID
   msLocationIEN = oSource.LocationIEN
   msClinicNCID = oSource.ClinicNCID
   mbDirty = oSource.IsDirty Or mbDirty
End Sub


Public Property Get RTF(ByVal sFacilityNCID As String) As String

On Error GoTo ErrorHandler

Dim sRTF As String
Dim lIdx As Long
Dim vCauseCodes As Variant

   If CausedByAccident Then
      sRTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" & _
               "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" & _
               "{\colortbl;\red0\green0\blue0;}" & _
               "\pard\f0\fs16" & _
            " Reported Injury/Accident:  Date " & Format(mdAccidentDate, "DD MMM YYYY") & ".  " & _
            "Reported Cause(s):  "
            
            vCauseCodes = Split(CauseCodeNCID, NCID_DELIMETER)
            For lIdx = 0 To UBound(vCauseCodes)
               
               sRTF = sRTF & GetWAMCauseCodeDesc(vCauseCodes(lIdx))
               
               If vCauseCodes(lIdx) = accAutoAccident Then
                  sRTF = sRTF & " (" & LocationName(sFacilityNCID) & ")"
               End If
               
               sRTF = sRTF & "; "
               
            Next
            
            'Trim off extra semi-colon (if any) and add period
            If Right(sRTF, 2) = "; " Then
               sRTF = Left(sRTF, Len(sRTF) - 2) & "."
            Else
               sRTF = sRTF & "."
            End If
                       
            'Close out the rtf section
            sRTF = sRTF & "\par\pard\tx3690\tx6480\pard }}"
            
   Else
      sRTF = vbNullString
   End If
   
   RTF = sRTF
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "RTF"
   
End Property
