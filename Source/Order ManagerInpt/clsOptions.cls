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
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const msDELIMITER_CRITERIA As String = "|"
Private mDefaultSettings As String
Private mStartDate      As Date
Private mEndDate        As Date
Private mDayMinus       As Integer

Public Enum ClinicOptions
    ThisClinic
    AllClinics
    SelectedClinics
End Enum

Public Enum ProviderOptions
    ThisProvider
    AllProviders
    SelectedProviders
End Enum

Public Enum Specialties
    AllSpecialtis
    SelectedSpecialtis
End Enum

Public Enum orderFilters
    orderDefault = 0
    allOrder = 1
    orderToday = 2
    orderTodayIncomORMinus = 4
    orderOneDay = 8
    orderDateRange = 16
    orderActiveORPend = 32
    orderComplete = 64
    orderDiscontinued = 128
    orderPreliminary = 256
    orderAddressed = 512
    orderUnAddressed = 1024
    orderDispensed = 2048
    allStatus = 4096
    orderSigned = 8192
    oWardSet = 16384
End Enum

Public Enum orderTypes
    hlLAB_ORDERS_CODE = 4
    hlRAD_ORDERS_CODE = 5
    hlRX_ORDERS_CODE = 9
    hlOTHER_ORDERS = 0
End Enum

Private msCriteriaBuf       As String
Public mClinicOption        As ClinicOptions
Public mProviderOption      As ProviderOptions
Public mcolSearchClinics    As Collection
Public mcolSearchProviders  As Collection
Public mCriteria            As orderFilters
Public mcolClinicProviders  As Collection
Public mDTOsMyClinics       As DTOs.Clinics
Public mAllOrderType        As Boolean
Public mOrderList           As Collection
Public mcolSearchWards      As Collection
Public mWardOption          As Specialties
Private Const ConfigKey As String = "settings"

Public Property Let mSelection(ByRef enmData As orderFilters)
  Call AddOrderSelectionCriteria(enmData)
End Property
Public Property Get mSelection() As orderFilters
    mSelection = mCriteria
End Property

Public Property Get mSelectionStatus() As String
  mSelectionStatus = msCriteriaBuf
End Property

Public Property Let mSelectionStatus(ByVal vData As String)
  mSelectionStatus = vData
End Property

Public Property Let StartDate(ByRef vData As Date)
    mStartDate = vData
End Property
Public Property Get StartDate() As Date
    StartDate = mStartDate
End Property
Public Property Let EndDate(ByVal vData As Date)
    mEndDate = vData
End Property
Public Property Get EndDate() As Date
    EndDate = mEndDate
End Property

Public Property Let DayMinus(ByVal vData As Integer)
    mDayMinus = vData
End Property

Public Property Get DayMinus() As Integer
    DayMinus = mDayMinus
End Property

Private Sub Class_Initialize()
    Call ClearOptions
End Sub

Public Sub ClearOptions()
    mEndDate = 0
    mStartDate = 0
    mAllOrderType = False
    msCriteriaBuf = vbNullString
    mCriteria = 0
    Set mOrderList = New Collection
    Set mcolSearchClinics = New Collection
    Set mcolSearchProviders = New Collection
    Set mcolSearchWards = New Collection
End Sub

Private Sub Class_Terminate()
    Set mOrderList = Nothing
    Set mcolSearchWards = Nothing
End Sub

Private Function UserIsProvider() As Boolean
    If goBB.User.HasPrivilegeEx(Priv_Orders_Laboratory_Orders, Priv_Sign) _
        Or goBB.User.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Sign) _
        Or goBB.User.HasPrivilegeEx(Priv_Orders_Radiology_Orders, Priv_Sign) _
        Or goBB.User.HasPrivilegeEx(Priv_Other_Therapies, Priv_Sign) Then
        UserIsProvider = True
    Else
        UserIsProvider = False
    End If
End Function


Friend Sub GetUserOptionsFromConfig()
    Dim sBuf As String
    On Error GoTo ErrHandler
       
    gbUserIsProvider = UserIsProvider
        
    '---get the user's default selections---
    If goBB.Config.GetValue(NCID_ORDER_MANAGER_INPATIENT_APPLICATION, "default settings", ConfigKey & goBB.Logon.UserNCID, sBuf) Then
        mDefaultSettings = sBuf
        If LenB(sBuf) = 0 Then
            Call SetInitSettingsToConfig
        End If
    Else
        Call SetInitSettingsToConfig
    End If
    
        
    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "GetUserOptionsFromConfig", "OrderMgt", vbExclamation
End Sub

Public Sub SetInitSettingsToConfig()
    On Error GoTo ErrHandler
          
    '--default settings variable does not exist, use predefined values as the default and save for this user--
    
    mDefaultSettings = AllClinics & ";;" & AllProviders
    
    StartDate = DateValue(Now) + TimeValue(#12:00:00 AM#)
    EndDate = DateValue(Now) + TimeValue(#11:59:00 PM#)
    
    mDefaultSettings = mDefaultSettings & ";;" & orderToday & "|" & allOrder & "|" & _
                        ";" & StartDate & ";" & EndDate
    
    goBB.Config.SetValue NCID_ORDER_MANAGER_INPATIENT_APPLICATION, goBB.Logon.UserNCID, "default settings", ConfigKey & goBB.Logon.UserNCID, mDefaultSettings
    goBB.Config.SaveChanges NCID_ORDER_MANAGER_INPATIENT_APPLICATION, goBB.Logon.UserNCID

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "SetInitSettingsToConfig", App.Title, vbExclamation

End Sub


Private Function GetProviderSetting() As String

    Dim bDidFirst As Boolean
    Dim objDTOsClinic As DTOs.Clinic
    Dim objDTOsProvider As DTOs.Provider
    Dim setting As String
    
    On Error GoTo ErrHandler
    
    bDidFirst = False
    setting = mClinicOption & ";"
    If mClinicOption = ThisClinic Then
        setting = setting & goBB.Logon.FacilityNCID
    Else
        For Each objDTOsClinic In mcolSearchClinics
          If Len(Trim$(objDTOsClinic.NCID)) > 0 Then
            If bDidFirst Then
                setting = setting & "," & objDTOsClinic.NCID
            Else
                bDidFirst = True
                setting = setting & objDTOsClinic.NCID
            End If
          End If
        Next objDTOsClinic
    End If
    
    setting = setting & ";" & mProviderOption & ";"
    bDidFirst = False
    If mProviderOption = ThisProvider Then
        setting = setting & goBB.Logon.UserName
    Else
        For Each objDTOsProvider In mcolSearchProviders
          If Len(Trim$(objDTOsProvider.NCID)) > 0 Then
            If bDidFirst Then
                setting = setting & "," & objDTOsProvider.NCID
            Else
                bDidFirst = True
                setting = setting & objDTOsProvider.NCID
            End If
          End If
        Next objDTOsProvider
    End If

    GetProviderSetting = setting
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".GetProviderSetting", App.Title, vbExclamation
End Function
Private Function GetOrderSetting() As String
    Dim i As Integer
    Dim setting As String
    
    On Error GoTo ErrHandler

    For i = 1 To mOrderList.Count
        setting = setting & mOrderList(i) & msDELIMITER_CRITERIA
    Next i
    
    GetOrderSetting = setting
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".GetOrderSetting", App.Title, vbExclamation
End Function
Public Sub SaveDefaultSettings()
    
    Dim prvdrSetting As String
    Dim orderSetting As String
    Dim wardSetting  As String
    
    On Error GoTo ErrHandler
    
    prvdrSetting = GetProviderSetting
    orderSetting = GetOrderSetting
    wardSetting = GetWardSetting
    
    mDefaultSettings = prvdrSetting & ";" & msCriteriaBuf & _
                        ";" & mStartDate & ";" & mEndDate & ";" & orderSetting & wardSetting & ";" & mDayMinus
                        
    goBB.Config.SetValue NCID_ORDER_MANAGER_INPATIENT_APPLICATION, goBB.Logon.UserNCID, _
                        "default settings", ConfigKey & goBB.Logon.UserNCID, mDefaultSettings
                        
    goBB.Config.SaveChanges NCID_ORDER_MANAGER_INPATIENT_APPLICATION, goBB.Logon.UserNCID
    
    
    Exit Sub
ErrHandler:
    '---show the VB error---
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".SaveDefaultSettings", App.Title, vbExclamation
End Sub

Public Sub SetDefaultSelections()
Dim aSettings() As String
    
On Error GoTo ErrHandler
        
    Call ClearOptions
    
    mClinicOption = ThisClinic
    If gbUserIsProvider Then
        mProviderOption = ThisProvider
    Else
        mProviderOption = AllProviders
    End If
       
    'DEFAULT SELECTIONS ARE SET
    aSettings = Split(mDefaultSettings, ";")
    
    mClinicOption = Val(aSettings(0))
    If LenB(aSettings(1)) > 0 Then
        LoadSearchClinics aSettings(1)
    End If
    
    mProviderOption = aSettings(2)
    If UBound(aSettings) < 3 Then
        Exit Sub
    End If
    If LenB(aSettings(3)) > 0 Then
        LoadSearchProviders aSettings(3)
    End If
    
    Call SetCriteriaFromString(aSettings(4))
    
    mStartDate = aSettings(5)
    If mStartDate = 0 Then
        mStartDate = DateValue(Now)
    End If
    
    mEndDate = aSettings(6)

    If UBound(aSettings) > 6 Then
        Call SetOrderTypeSetting(aSettings(7))
    Else
        mAllOrderType = True
    End If
    
    If UBound(aSettings) > 7 Then
        Call SetWardSetting(aSettings(8))
    End If
    
    If UBound(aSettings) > 8 Then
        mDayMinus = Val(aSettings(9))  'Get minus days
    End If
    
    Exit Sub
    
ErrHandler:
    Call SetInitSettingsToConfig
    
End Sub
Private Sub SetOrderTypeSetting(ByVal orderTypes As String)

    On Error GoTo ErrHandler
    Dim oTypeArray() As String
    Dim i As Integer
    Dim iUpper As Integer
       
    oTypeArray() = Split(orderTypes, msDELIMITER_CRITERIA)
      
    If LenB(Trim$(orderTypes)) = 0 Then
        mAllOrderType = True
    Else
        iUpper = UBound(oTypeArray)
        For i = 0 To iUpper
            If LenB(oTypeArray(i)) > 0 Then
                mOrderList.Add Trim$(oTypeArray(i))
            End If
        Next i
    End If
    
    Exit Sub
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, ".SetCriteriaFromString", "OrderMgt", vbExclamation)
End Sub
Private Sub SetCriteriaFromString(ByVal sCriteria As String)
    On Error GoTo ErrHandler
    Dim sCriteriaArray() As String
    Dim i As Integer
    Dim iUpper As Integer
    Dim lCriteria As Long
       
    sCriteriaArray() = Split(sCriteria, msDELIMITER_CRITERIA)
      
    If LenB(Trim$(sCriteria)) = 0 Then
        mSelection = orderToday
    Else
        iUpper = UBound(sCriteriaArray)
        For i = 0 To iUpper
            lCriteria = CLng(Val((Trim$(sCriteriaArray(i)))))
            If Len(sCriteriaArray(i)) Then
                mSelection = lCriteria
            End If
        Next i
    End If
    
    Exit Sub
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, ".SetCriteriaFromString", "OrderMgt", vbExclamation)

End Sub
Public Function AddOrderSelectionCriteria(ByRef enmData As orderFilters) As Boolean
  Dim sData As String
  
  On Error GoTo ErrHandler
  
  sData = CStr(enmData)
  
    If Not mCriteria And enmData Then
        mCriteria = mCriteria Or enmData
        msCriteriaBuf = msCriteriaBuf & sData & msDELIMITER_CRITERIA
    
      End If
    
      Exit Function

ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "frmOrderMgt.AddOrderSelectionCriteria", "Order Manager", vbExclamation

End Function

Private Sub LoadSearchClinics(ByVal Clinics As String)
Dim aNCIDs() As String
Dim i As Integer
        
  On Error Resume Next
  
  aNCIDs = Split(Clinics, ",")
  
  For i = LBound(aNCIDs) To UBound(aNCIDs)
    If Len(Trim$(aNCIDs(i))) > 0 Then
        mcolSearchClinics.Add mDTOsMyClinics(aNCIDs(i)), aNCIDs(i)
    End If
  Next

End Sub
Private Sub LoadSearchProviders(ByVal Providers As String)
Dim aNCIDs() As String
Dim i As Integer
Dim colProviders As DTOs.Providers
Dim objDTOsProvider As DTOs.Provider
        
  On Error GoTo ErrHandler
  
  aNCIDs = Split(Providers, ",")
  
  For i = LBound(aNCIDs) To UBound(aNCIDs)
      For Each colProviders In mcolClinicProviders
          For Each objDTOsProvider In colProviders
              If objDTOsProvider.NCID = aNCIDs(i) Then
                  mcolSearchProviders.Add objDTOsProvider, objDTOsProvider.NCID
                  GoTo NextNCID
              End If
          Next objDTOsProvider
      Next colProviders

NextNCID:
  Next

  Set objDTOsProvider = Nothing
      
  Exit Sub

ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "frmOrderMgt.LoadSearchProviders", "OrderMgt", vbExclamation

End Sub
Public Function GetSelectionLabel() As String
    Dim sLabelWho                 As String
    Dim sLabelWhere               As String
    Dim sLabelWhen                As String
    Dim sLabelWhat                As String
    
    sLabelWhere = goBB.Logon.FacilityName
   
    sLabelWhen = DecodeWhenString
    sLabelWho = DecodeWhoString
    sLabelWhat = DecodeStatusString

    If InStr(1, sLabelWhen, sLabelWhat) = 0 Then
        GetSelectionLabel = sLabelWho & " in " & sLabelWhere & " for " & sLabelWhen & sLabelWhat & "."
    Else
        GetSelectionLabel = sLabelWho & " in " & sLabelWhere & " for " & sLabelWhen & "."
    End If

End Function
Private Function DecodeStatusString() As String
  Dim sLabelWhat As String
  
  sLabelWhat = Empty
        
    If goBB.Options.mCriteria And allOrder Then
        sLabelWhat = "All Status"
        If mCriteria And orderSigned Then
            sLabelWhat = sLabelWhat & " or Unsigned"
        End If
        DecodeStatusString = sLabelWhat
    Else
        If mCriteria And orderPreliminary Then
            sLabelWhat = sLabelWhat & " or Preliminary"
        End If
        
        If mCriteria And orderDispensed Then
            sLabelWhat = sLabelWhat & " or Dispensed"
        End If
        
        If mCriteria And orderAddressed Then
            sLabelWhat = sLabelWhat & " or Addressed"
        End If
        
        If mCriteria And orderComplete Then
            sLabelWhat = sLabelWhat & " or Complete"
        End If
        
        If mCriteria And orderUnAddressed Then
            sLabelWhat = sLabelWhat & " or Unaddressed"
        End If
        'new code for inpatient
        If mCriteria And orderActiveORPend Then
            sLabelWhat = sLabelWhat & " or Active/Pending"
        End If
        
        If mCriteria And orderDiscontinued Then
            sLabelWhat = sLabelWhat & " or Discontinued"
        End If
        
        If mCriteria And orderSigned Then
            sLabelWhat = sLabelWhat & " or Unsigned"
        End If

        DecodeStatusString = Replace(sLabelWhat, " or ", vbNullString, , 1)
   End If

        
End Function
Private Function DecodeWhoString() As String
Dim sLabelWho                 As String

  On Error GoTo ErrHandler
  
  Select Case mProviderOption
    Case ThisProvider
        sLabelWho = "My Orders"
    
    Case AllProviders
        sLabelWho = "Orders for All Providers"
    
    Case SelectedProviders
        If mcolSearchProviders.Count = 0 Then
            mProviderOption = ThisProvider
            sLabelWho = "My Orders"
        ElseIf mcolSearchProviders.Count = 1 Then
            sLabelWho = "Orders for " & mcolSearchProviders.Item(1).Name
        Else
            sLabelWho = "Orders for " & mcolSearchProviders.Item(1).Name & ", et al.,"
        End If
  End Select
  
  DecodeWhoString = sLabelWho
  
  Exit Function
  
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOption.DecodeWhoString", "OrderMgt", vbExclamation

End Function
Private Function DecodeWhenString() As String
Dim sLabelWhen                As String
  
    On Error GoTo ErrHandler

    If mCriteria And orderToday Then
        sLabelWhen = "Today only "
    ElseIf mCriteria And orderTodayIncomORMinus Then
        sLabelWhen = "Today Minus " & mDayMinus & " days "
    
    ElseIf mEndDate = vbEmpty _
      Or (mCriteria And orderOneDay) Then
        If mStartDate = vbEmpty Or mStartDate = #12/31/1899# Then
            mStartDate = Now
        End If
        sLabelWhen = Format$(mStartDate, "dd MMM yyyy") & " "
    Else
        sLabelWhen = Format$(mStartDate, "dd MMM yyyy HHnn") & " thru " & Format$(mEndDate, "dd MMM yyyy HHnn") & " "
    End If

    DecodeWhenString = sLabelWhen
    
  Exit Function
  
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOption.DecodeWhenString", "OrderMgt", vbExclamation
End Function

Public Function CheckUserAccassPriv() As Boolean
Dim objDTOsClinic As DTOs.Clinic
On Error GoTo ErrHandler

    Set mDTOsMyClinics = goBB.ClinicOps.GetUserClinics(goBB.Logon.FacilityNCID, goBB.Logon.UserNCID)
    If mDTOsMyClinics.Count < 1 Then
        CheckUserAccassPriv = False
        Exit Function
    End If
    
    'SCR 174211 always reload the mcolClinicProviders collection when loading the form
    Set mcolClinicProviders = Nothing
    
    Set mcolClinicProviders = New Collection
    For Each objDTOsClinic In mDTOsMyClinics
        mcolClinicProviders.Add goBB.ClinicOps.GetClinicProviders(objDTOsClinic.NCID), objDTOsClinic.NCID
    Next objDTOsClinic
    Set objDTOsClinic = Nothing

    
    CheckUserAccassPriv = True

  Exit Function
  
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOption.CheckUserAccassPriv", "OrderMgt", vbExclamation
End Function

Public Sub GetClinicans()
    Dim objDTOsProvider           As DTOs.Provider
    On Error GoTo ErrHandler
    
    With gobjOrderMgtInpt
        If mProviderOption = ThisProvider Then
            .CriteriaAddClinician goBB.Logon.UserNCID
        ElseIf mProviderOption = SelectedProviders Then
            For Each objDTOsProvider In mcolSearchProviders
                .CriteriaAddClinician objDTOsProvider.NCID
            Next objDTOsProvider
        End If
    End With
    
    Set objDTOsProvider = Nothing
    Exit Sub
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOption.GetClinicans", "OrderMgt", vbExclamation)
End Sub

Public Sub GetClinics()
    Dim objDTOsClinic           As DTOs.Clinic
    On Error GoTo ErrHandler
    
    With gobjOrderMgtInpt
        If mClinicOption = ThisClinic Then
            .CriteriaAddClinic goBB.Logon.NursingDivisionNCID
        ElseIf mClinicOption = SelectedClinics Then
            If mcolSearchClinics.Count = 0 Then
                For Each objDTOsClinic In mDTOsMyClinics
                  If Len(Trim$(objDTOsClinic.NCID)) > 0 Then
                        .CriteriaAddClinic objDTOsClinic.NCID
                  End If
                Next objDTOsClinic
            Else
                For Each objDTOsClinic In mcolSearchClinics
                  If Len(Trim$(objDTOsClinic.NCID)) > 0 Then
                        .CriteriaAddClinic objDTOsClinic.NCID
                  End If
                Next objDTOsClinic
            End If
        End If
    End With
    
    Set objDTOsClinic = Nothing
    Exit Sub
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOption.GetClinics", "OrderMgt", vbExclamation)
End Sub

Private Sub SetWardSetting(ByVal oWards As String)

    On Error GoTo ErrHandler
    Dim oTypeArray() As String
    Dim i As Integer
    Dim iUpper As Integer
    Dim oWard As cWard
    
    oTypeArray() = Split(oWards, msDELIMITER_CRITERIA)
      
    If LenB(Trim$(oWards)) > 0 Then
        iUpper = UBound(oTypeArray)
        mWardOption = oTypeArray(0)
        For i = 1 To iUpper
            If LenB(oTypeArray(i)) > 0 Then
                Set oWard = New cWard
                oWard.WardNCID = Trim$(oTypeArray(i))
                mcolSearchWards.Add oWard
            End If
        Next i
    End If
    
    Exit Sub
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOption.SetWardSetting", "OrderMgt", vbExclamation)
End Sub

Private Function GetWardSetting() As String
    Dim i     As Integer
    Dim oWard As cWard
    Dim setting As String
    
    On Error GoTo ErrHandler

    setting = ";" & mWardOption & msDELIMITER_CRITERIA
    
    For i = 1 To mcolSearchWards.Count
        Set oWard = mcolSearchWards(i)
        setting = setting & oWard.WardNCID & msDELIMITER_CRITERIA
    Next i
    
    GetWardSetting = setting
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOption.GetWardSetting", App.Title, vbExclamation
End Function

Public Sub ResetSelectionStatus(ByRef enmData As orderFilters)
    Dim arrStatusCriteria() As String
    Dim eData As orderFilters
    On Error GoTo ErrHandler
    
    arrStatusCriteria = Split(msCriteriaBuf, "|")
    mCriteria = 0
    msCriteriaBuf = vbNullString
    eData = arrStatusCriteria(0)
    AddOrderSelectionCriteria (eData)
    AddOrderSelectionCriteria (enmData)

    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOption.ResetSelectionStatus", App.Title, vbExclamation
End Sub
