VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XlateNCID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Type Lexicon
  sName As String
  sNCID As String
End Type
'Dim objDas As DasCtrl
Dim mRoute() As Lexicon
Dim mSite() As Lexicon
Dim mStatus() As Lexicon
Dim mAdminLocation() As Lexicon
Dim mSeqNumber() As Lexicon
Dim mVaccine() As Lexicon
Dim mUnits() As Lexicon
Dim mClinician() As Lexicon
'Dim mClinician2() As Lexicon
Dim mManufacturer() As Lexicon
Dim mCodedComment() As Lexicon

Public Sub Setup(sEnterpriseNCID As String)
'WriteMessage "about to start getproviders", "objXLate.setup"
'         'Fill list of Providers!!!!
'        modCISM.GetProvidersFromCDR mClinician
'WriteMessage "just gotproviders", "objXLate.setup"
'
        '--- load pick list of locations/facilities
        If Not loadArray(sEnterpriseNCID, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mAdminLocation) Then
            'Exit Sub
        End If
        
        '--- load pick list of routes ---
        If Not loadArray(NCID_ROUTE_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mRoute) Then
            'Exit Sub
        End If
        
        '---load pick list of units---
        If Not loadArray(NCID_UNITS_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mUnits) Then
            'Exit Sub
        End If
        
        '---load pick list of manufacturers---
        'If Not loadArray(NCID_MANUFACTURER_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mManufacturer) Then
        If Not loadArray(NCID_MANUFACTURER_DOMAIN, "54366", NCID_HAS_MEMBER, mManufacturer) Then
            '54366 is the representation NCID for mfgr codes!
            'Exit Sub
        End If
        
        If Not loadArray(NCID_MAJOR_VACCINE_CLASS_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mVaccine) Then
            'Exit Sub
        End If
        
        If Not loadArray(sEnterpriseNCID, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mAdminLocation) Then
            'Exit Sub
        End If
        '---load pick list of sequence numbers---
        'If Not loadArray(NCID_IMMUNIZATION_SEQ_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mSeqNumber) Then
        '    Exit Sub
        'End If
        
        If Not loadArray(NCID_IMMUN_CODED_COMMENTS_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mCodedComment) Then
            'Exit Sub
        End If
        
        If Not loadArray(NCID_MED_ADMIN_EVENT_STATUS_DOMAIN, NCID_3M_DEFAULT_CONTEXT, NCID_HAS_MEMBER, mStatus) Then
            'Exit Sub
        End If
  

 
End Sub

Public Function RouteName(sRouteNCID As String) As String
Select Case sRouteNCID
    Case "67590"
      RouteName = "IM"
    Case "67610"
      RouteName = "SC"
    Case "67595"
      RouteName = "ID"
    Case "67605"
      RouteName = "Oral"
    Case Else
      RouteName = ""
End Select
'RouteName = GetName(mRoute, sRouteNCID)
End Function
Public Function RouteNCID(sRouteName As String) As String
'RouteNCID = GetNCID(mRoute, sRouteName)
Select Case sRouteName
  Case "IM", "Intramuscular"
    RouteNCID = "67590"
  Case "SC", "subcutaneous"
    RouteNCID = "67610"
  Case "ID", "intradermal"
    RouteNCID = "67595"
  Case "Oral"
    RouteNCID = "67605"
  Case Else
    RouteNCID = ""
End Select
End Function

Public Function SiteName(sRouteNCID As String) As String
    
Select Case sRouteNCID
    Case NCID_ORAL
        SiteName = "Oral"
    Case NCID_LEFT_ARM
        SiteName = "Arm, Left"
    Case NCID_RIGHT_ARM
        SiteName = "Arm, Right"
    Case NCID_LEFT_THIGH
        SiteName = "Thigh, Left"
    Case NCID_RIGHT_THIGH
        SiteName = "Thigh, Right"
    Case Else
        SiteName = ""
End Select

End Function

Public Function SiteNCID(sRouteName As String) As String
    Select Case sRouteName
        Case "Oral"
          SiteNCID = NCID_ORAL
        Case "Arm, Left", "Left Arm"
          SiteNCID = NCID_LEFT_ARM
        Case "Arm, Right", "Right Arm"
          SiteNCID = NCID_RIGHT_ARM
        Case "Thigh, Left", "Left Thigh"
          SiteNCID = NCID_LEFT_THIGH
        Case "Thigh, Right", "Right Thigh"
          SiteNCID = NCID_RIGHT_THIGH
        Case Else
          SiteNCID = ""
    End Select

End Function

Public Function StatusName(sStatusNCID As String) As String
StatusName = GetName(mStatus, sStatusNCID)
End Function

Public Function StatusNCID(sStatusName As String) As String
StatusNCID = GetNCID(mStatus, sStatusName)
End Function

Public Function AdminLocationName(sAdminLocationNCID As String) As String
AdminLocationName = GetName(mAdminLocation, sAdminLocationNCID)
End Function

Public Function AdminLocationNCID(sAdminLocationName As String) As String
AdminLocationNCID = GetNCID(mAdminLocation, sAdminLocationName)
End Function

Public Function SeqNumberName(sSeqNumberNCID As String) As String
'SeqNumberName = GetName(mSeqNumber, sSeqNumberNCID)
Select Case sSeqNumberNCID
    Case "162061"
      sSeqNumberName = "1"
    Case "162062"
      sSeqNumberName = "2"
    Case "162063"
      sSeqNumberName = "3"
    Case "162064"
      sSeqNumberName = "4"
    Case "14500056"
      sSeqNumberName = "5"
    Case "14500057"
      sSeqNumberName = "6"
    Case "162066"
      sSeqNumberName = "Booster"
    Case "162065"
      sSeqNumberName = "Primary Series Complete"
    Case Else
      sSeqNumberName = ""
End Select
  
End Function

Public Function SeqNumberNCID(sSeqNumberName As String) As String
'SeqNumberNCID = GetNCID(mSeqNumber, sSeqNumberName)
  Select Case Trim(sSeqNumberName)
    Case "1", "# 1", "0"
      SeqNumberNCID = "162061"
    Case "2", "# 2"
      SeqNumberNCID = "162062"
    Case "3", "# 3"
      SeqNumberNCID = "162063"
    Case "4", "# 4"
      SeqNumberNCID = "162064"
    Case "5", "# 5", "Dose - 5"
      SeqNumberNCID = "14500056"
    Case "6", "# 6", "Dose - 6"
      SeqNumberNCID = "14500057"
    Case "Booster"
      SeqNumberNCID = "162066"
    Case "Primary Series Complete"
      SeqNumberNCID = "162065"
    Case Else
      SeqNumberNCID = "0"
    End Select
    
End Function

Public Function VaccineName(sVaccineNCID As String) As String
VaccineName = GetName(mVaccine, sVaccineNCID)
End Function

Public Function VaccineNCID(sVaccineName As String) As String
VaccineNCID = GetNCID(mVaccine, sVaccineName)
End Function

Public Function UnitsName(sUnitsNCID As String) As String
UnitsName = GetName(mUnits, sUnitsNCID)
End Function

Public Function UnitsNCID(sUnitsName As String) As String
UnitsNCID = GetNCID(mUnits, sUnitsName)
End Function

Public Function ClinicianName(sClinicianNCID As String) As String

End Function
Public Function ClinicianNCID(Optional sClinicianSSN As String, Optional sClinicianName As String) As String


End Function


Public Function ManufacturerName(sManufacturerNCID As String) As String
ManufacturerName = GetName(mManufacturer, sManufacturerNCID)
End Function
Public Function ManufacturerNCID(sManufacturerName As String) As String
ManufacturerNCID = GetNCID(mManufacturer, sManufacturerName)
End Function
Public Function CommentsName(sCodedCommentsNCID As String) As String
CommentsName = GetName(mCodedComment, sCodedCommentsNCID)
End Function
'****************************************************************************************************************************
'  Function Name: loadArray
'
'  Description: This function loads the pick-list for the passed in criteria.
'
'  Parameters:
'    vsDomainNCID - Contains the "domain" for which to build the pick-list.
'    vsContextNCID - Contains the "context" to use in building the pick-list.
'    vsRelationshipNCID - Contains the "relationship" for which to build the pick-list.
'    rcboPickList - Upon return will contain the pick-list items the user can choose from.
'    vbAddEmptyItem - TRUE if an "empty" item needs to be added to the pick-list (used for
'                     "clearing" a selected item, FALSE otherwise.
'
'  Return Value:
'    Returns TRUE if the pick-list was loaded successfully.
'    Returns FALSE if an error occurred in loading the pick-list.
'****************************************************************************************************************************
Private Function loadArray(ByVal vsDomainNCID As String, _
                                  ByVal vsContextNCID As String, _
                                  ByVal vsRelationshipNCID As String, _
                                  myArray() As Lexicon) As Boolean
    Dim lNumItems As Long
    Dim objPickList As PickListCtrl
    Dim i As Integer

    On Error GoTo ErrHandler
    
    '---set return value---
    loadArray = False
    
    '---clear out current items---
    'rcboPickList.Clear

    '---create a "pick-list" object---
    Set objPickList = New PickListCtrl
    
    '---set criteria for "loading" the pick-list ---
    objPickList.Domain = vsDomainNCID
    objPickList.Context = vsContextNCID
    objPickList.Relationship = vsRelationshipNCID
    objPickList.PrefOnlyFlag = True
    objPickList.Load

    '---initialize variables---
    lNumItems = objPickList.ListItemCnt

    ReDim myArray(lNumItems)
    If lNumItems > 0 Then
        '---load the "pick-list" items---
        For i = 1 To lNumItems
            '---store the "representation" and the "NCID" of each pick-list item in the passed in combo-box---
            myArray(i).sName = objPickList.ListItem(i).Representation
            myArray(i).sNCID = objPickList.ListItem(i).Concept
        Next i
        
               '---set return value---
        loadArray = True
    End If
    
    '---release reference to object---
    Set objPickList = Nothing
    
    Exit Function

ErrHandler:
    '---show the VB error---
    ReDim myArray(0)
    loadArray = False
    '---release reference to object---
    Set objPickList = Nothing

    Exit Function
End Function




Private Function GetNCID(myArray() As Lexicon, MyName As String) As String
Dim i As Long
Dim tmpname As String
  tmpname = UCase(Trim(MyName))
  GetNCID = ""
  For i = LBound(myArray) To UBound(myArray)
    If UCase(myArray(i).sName) = tmpname Then
      GetNCID = myArray(i).sNCID
      Exit For
    End If
  Next i

End Function

Private Function GetName(myArray() As Lexicon, MyNCID As String) As String
Dim i As Long
  GetName = ""
  For i = LBound(myArray) To UBound(myArray)
    If myArray(i).sNCID = MyNCID Then
      GetName = myArray(i).sName
      Exit For
    End If
  Next i
End Function

Public Function GetVaccines(Names() As String, ncids() As String)
Dim i As Integer
Dim j As Integer

  i = UBound(mVaccine)
  ReDim Names(i)
  ReDim ncids(i)
  For j = 1 To i
    Names(j) = mVaccine(j).sName
    ncids(j) = mVaccine(j).sNCID
  Next j

End Function

Public Function GetAllergens(Names() As String, TypeNCIDs() As String, ncids() As String, objconfig As Object)
  Load frmAllergens
  frmAllergens.GetTheAllergens Names, TypeNCIDs, ncids, objconfig
  Unload frmAllergens
End Function

Public Function GetRoutes(Names() As String, ncids() As String)
Dim i As Integer
Dim j As Integer

  i = UBound(mRoute)
  ReDim Names(i)
  ReDim ncids(i)
  For j = 1 To i
    Names(j) = mRoute(j).sName
    ncids(j) = mRoute(j).sNCID
  Next j


End Function
Public Function GetProviders(Names() As String, ncids() As String)
Dim i As Integer
Dim j As Integer

  i = UBound(mClinician)
  ReDim Names(i)
  ReDim ncids(i)
  For j = 1 To i
    Names(j) = mClinician(j).sName
    ncids(j) = mClinician(j).sNCID
  Next j



End Function
Public Function GetLocations(Names() As String, ncids() As String)
Dim i As Integer
Dim j As Integer

  i = UBound(mAdminLocation)
  ReDim Names(i)
  ReDim ncids(i)
  For j = 1 To i
    Names(j) = mAdminLocation(j).sName
    ncids(j) = mAdminLocation(j).sNCID
  Next j


End Function
