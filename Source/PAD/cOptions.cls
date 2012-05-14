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
Attribute VB_Name = "cOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Allspecialties As Boolean
Public SpecialtyNCIDs As Collection   ' of ncids for selected specialties
Public ActiveSpecialties As Boolean

Public AllProviders   As Boolean
Public JustMe         As Boolean
Public ProviderNCIDs  As Collection   ' of provider ncids

Public StatusSelection As Integer  ' sum of inpatient statuses "ANY"=0, "CURRENT"=2,"DISHARGED"=4
Public ShowAllBeds    As Boolean
Public ShowAllInpatients As Boolean
Public TodayOnly      As Boolean
Public TodayAndInpatient As Boolean
Public OneDate        As Boolean
Public SelectedDate   As Date
Public DateRange      As Boolean
Public DateRangeFrom  As Date
Public DateRangeTo    As Date

Private objConfig     As MMMHISConfiguration.Config
Public ObjectStatus   As ObjectState_
Public CurrSQL        As String         '__ the sql string created in loadpatients

Public Property Get Description() As String      '-- the text describing the options selected
        Description = BuildDescription
End Property
Public Property Set Config(RHV As MMMHISConfiguration.Config)
    Set objConfig = RHV
End Property

Private Sub Class_Initialize()
    Call SetDefaults
End Sub

Public Sub SetDefaults()
    '-- UserOptions is a '^'delimited string in this order:
'0 Allspecialties
'1 SpecialtyNCIDs (a |-delimited string of specialty ncids)
'2 AllProviders
'3 JustMe
'4 ProviderNCIDs  (a |-delimited string of provider ncids)

'5 StatusSelection  ' sum of inpatient statuses "ANY"=2, "INPATIENT"=4,"DISHARGED"=8
'6 ShowAllBeds
'7 TodayOnly
'8 TodayAndInpatient
'9 OneDate
'10 SelectedDate
'11 DateRange
'12 DateRangeFrom
'13 DateRangeTo
'14 ActiveSpecialties

    Set SpecialtyNCIDs = New Collection
    Set ProviderNCIDs = New Collection
    Allspecialties = True
    'SCR 103827 default to Me
    AllProviders = False
    JustMe = True
    
    StatusSelection = os_inpatient
    ShowAllBeds = False
    ShowAllInpatients = False
    TodayOnly = False
    TodayAndInpatient = True
    OneDate = False
    SelectedDate = 0
    DateRange = False
    DateRangeFrom = 0
    DateRangeTo = 0
    ActiveSpecialties = True
    
'    If UserOptions = vbnullstring Then
'        '-- go with defaults
'        UserOptions = "-1^^0^-1^^4^0^0^-1^0^^0^^^"
'    End If

    
End Sub

Private Sub Class_Terminate()
    Set SpecialtyNCIDs = Nothing
    Set ProviderNCIDs = Nothing
End Sub

Public Function ParseUserOptions(UserOptions As String) As Boolean
    
    Dim opt() As String
    Dim NCIDS() As String
    Dim i As Long
    
On Error GoTo ErrHandler
    
    opt = Split(UserOptions, "^")
    If UBound(opt) < 15 Then
        '__ again, take the default
        'MsgBox "Unable to retrieve User Options.  Default options were loaded", vbInformation, "User Options Not Available"
        SetDefaults
        ParseUserOptions = True
        Exit Function
    End If
        Allspecialties = opt(0)
        If opt(1) <> vbNullString Then
            NCIDS = Split(opt(1), "|")
            For i = 0 To UBound(NCIDS)
                If NCIDS(i) <> vbNullString Then
                    SpecialtyNCIDs.Add NCIDS(i)
                End If
            Next i
        End If
        AllProviders = opt(2)
        JustMe = opt(3)
        If opt(4) <> vbNullString Then
            NCIDS = Split(opt(4), "|")
            For i = 0 To UBound(NCIDS)
                If NCIDS(i) <> vbNullString Then
                    ProviderNCIDs.Add NCIDS(i)
                End If
            Next i
        End If
        StatusSelection = opt(5)
        ShowAllBeds = opt(6)
        ShowAllInpatients = opt(7)
        TodayOnly = opt(8)
        TodayAndInpatient = opt(9)
        OneDate = opt(10)
        SelectedDate = opt(11)
        DateRange = opt(12)
        DateRangeFrom = opt(13)
        DateRangeTo = opt(14)
'-- added 12.10.03
    If UBound(opt) = 15 Then
        ActiveSpecialties = opt(15)
    End If
    
    Exit Function
ErrHandler:
    SetDefaults
    ParseUserOptions = True
End Function

Friend Sub GetUserOptions()
'__ get the user options from the osh.  If none found, setup the default
    Dim UserOptions As String
    
    gobjconfig.GetValue NCID_INPATIENT_APPLICATION, "INPATIENT", "SELECTED_OPTIONS", UserOptions
    '-- UserOptions is a '^'delimited string in this order:
'0 Allspecialties
'1 SpecialtyNCIDs (a |-delimited string of specialty ncids)
'2 AllProviders
'3 JustMe
'4 ProviderNCIDs  (a |-delimited string of provider ncids)

'5 StatusSelection  ' sum of inpatient statuses "ANY"=2, "INPATIENT"=4,"DISHARGED"=8
'6 ShowAllBeds
'7 ShowAllInpatients
'8 TodayOnly
'9 TodayAndInpatient
'10 OneDate
'11 SelectedDate
'12 DateRange
'13 DateRangeFrom
'14 DateRangeTo
'15 activespecialties  only
    Call ParseUserOptions(UserOptions)
    
End Sub

Private Function BuildDescription() As String
    Dim Description As String
    Description = vbNullString
'__ type of patient
    If TodayAndInpatient Or (StatusSelection And os_inpatient) Then
        Description = "Current inpatient records"
    ElseIf (StatusSelection And os_discharged) Then
        Description = "Discharged patient records"
    Else
        Description = "Any patient records"
    End If
'__ attending
    If JustMe Then
        Description = Description & " with " & gobjLogon.StoringUserName & " as attending provider"
    ElseIf AllProviders Then
        Description = Description    '- all providers
    Else
        Description = Description & " for selected attending providers"
    End If
'__ dates
    If TodayOnly Then
        Description = Description & " admitted today"
    ElseIf TodayAndInpatient Then  '--- SCR 72589   Sherry Wang   7/5/2005
        '--- SCR 81979   Sherry Wang  12/2/2005
        Description = Description & " A&&D today plus inpatient"
    ElseIf OneDate Then
        Description = Description & " admitted " & Format$(SelectedDate, "DD-MMM-YYYY")
    ElseIf DateRange Then
        Description = Description & " admitted between " & Format$(DateRangeFrom, "DD-MMM-YYYY hh:nn") & " and  " & Format$(DateRangeTo, "DD-MMM-YYYY hh:nn")
    End If
'__ wards
    If Allspecialties Then
        If ActiveSpecialties Then
            Description = Description & " in any active specialty"
        Else
            Description = Description & " in any specialty"
        End If
    Else
        Description = Description & " in selected specialties"
    End If
'__ bed status
    If ShowAllBeds Then
        Description = Description & " (showing all beds)."
    Else
        Description = Description & "."
    End If
    BuildDescription = Description
End Function

Friend Function SaveUserOptions() As Boolean
    Dim UserOptions As String
    Dim vItem As Variant
On Error GoTo ErrHandler

    '__ build the string
        UserOptions = Allspecialties & "^"
        For Each vItem In SpecialtyNCIDs
            UserOptions = UserOptions & vItem & "|"
        Next vItem
        '--drop the trailing pipe if one exists
        If Right$(UserOptions, 1) = "|" Then
            UserOptions = Left$(UserOptions, Len(UserOptions) - 1)
        End If
        UserOptions = UserOptions & "^" & AllProviders & "^" _
                                        & JustMe & "^"
        For Each vItem In ProviderNCIDs
            UserOptions = UserOptions & vItem & "|"
        Next vItem
        '--drop the trailing pipe
        If Right$(UserOptions, 1) = "|" Then
            UserOptions = Left$(UserOptions, Len(UserOptions) - 1)
        End If
        
        UserOptions = UserOptions & "^" & StatusSelection & "^" _
                                        & ShowAllBeds & "^" _
                                        & ShowAllInpatients & "^" _
                                        & TodayOnly & "^" _
                                        & TodayAndInpatient & "^" _
                                        & OneDate & "^" _
                                        & SelectedDate & "^" _
                                        & DateRange & "^" _
                                        & DateRangeFrom & "^" _
                                        & DateRangeTo & "^" _
                                        & ActiveSpecialties
        '-- sort order
        UserOptions = UserOptions
    gobjconfig.SetValue NCID_INPATIENT_APPLICATION, gobjLogon.UserNCID, "INPATIENT", "SELECTED_OPTIONS", UserOptions
    gobjconfig.SaveChanges NCID_INPATIENT_APPLICATION, gobjLogon.UserNCID
    
    SaveUserOptions = True
    Exit Function
ErrHandler:
    MsgBxARMd "Unable to save user options", vbInformation, "Save User Options"
End Function


