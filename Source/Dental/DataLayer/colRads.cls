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
Attribute VB_Name = "colRads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsRad"
Attribute VB_Ext_KEY = "Member0" ,"clsRad"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlPatientUnitID As Long
Private moClinicOps As ClinicOps

'These are used for creating rads outside of chcsii - Image Store SCP
Private mlExtRadOrderID As Long
Private msExtLocation As String
Private meExtRadType As eRadType
Private mlExtPatientUnitID As Long
Private mlExtRequestedByNCID As Long
Private mlExtFacilityNCID As Long

Public Function AddLite(Status As eRadStatus, Version As Long) As clsRad

'This is used outside of CHCSII and ExternalInit should be called first

On Error GoTo ErrHnd

    'create a new object
    Dim objNewMember As clsRad
    Set objNewMember = New clsRad
    Dim oRs As Recordset

    objNewMember.RadID = goconn.GetNextID(esDENTAL_IMAGING_RADS_SEQ)
    
    With objNewMember
        'Set the required properties
        .RadOrderID = mlExtRadOrderID
        .Location = msExtLocation
        .RadType = meExtRadType
        .PatientUnitID = mlExtPatientUnitID
        
        Set cwobjCWComm = New MMMHISComm.Comm
        'GMT time
        .RadDate = cwobjCWComm.GetDateTime
        .RequestedByNCID = mlExtRequestedByNCID
        .FacilityNCID = mlExtFacilityNCID
        .StatusID = Status
        .Version = Version
        .DirtyFlag = dfAdd
    End With
    
    'Add it to the beginning of the collection
    If mCol.Count > 0 Then
        mCol.Add objNewMember, CStr(objNewMember.RadID), 1
    Else
        mCol.Add objNewMember, CStr(objNewMember.RadID)
    End If

    'return the object created
    Set AddLite = objNewMember
    Set objNewMember = Nothing

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    Resume
End Function


Private Function CompareArrays(Search As String, Find As String) As Boolean

'This function converts the strings into arrays and looks for elements from the Find array that are in the Search array
'If one index matches, the it returns true

Dim arrSearch() As String
Dim arrFind() As String
Dim iOuter As Long
Dim iInner As Long

On Error GoTo ErrHnd

    arrSearch = Split(Search, ",")
    arrFind = Split(Find, ",")
    
    For iOuter = 0 To UBound(arrFind)
        For iInner = 0 To UBound(arrSearch)
            If Trim(arrFind(iOuter)) = Trim(arrSearch(iInner)) Then
                CompareArrays = True
                Exit For
            End If
        Next
        If CompareArrays = True Then
            Exit For
        End If
    Next

Exit Function

ErrHnd:
    MsgBox "Error in CompareArrays: " & Err.Description
    
End Function


Public Sub ExternalInit(RadOrderID As Long, Location As String, RadType As eRadType, PatientUnitID As Long, RequestedByNCID As Long, FacilityNCID As Long)

'This is called from clsRadOrder
'This is used to save variables that are common to the rad order when creating rads outside of chcsii
'This is used in AddLite

    mlExtRadOrderID = RadOrderID
    msExtLocation = Location
    meExtRadType = RadType
    mlExtPatientUnitID = PatientUnitID
    mlExtRequestedByNCID = RequestedByNCID
    mlExtFacilityNCID = FacilityNCID

End Sub

Public Function Find(Filter As tRadFilter) As colReadOnlyRads

'This loops through all the objects in the collection and finds the matches based on the properties set in Filter

Dim oRad As clsRad
Dim oReadOnlyRads As colReadOnlyRads
Dim bMatched As Boolean
    
    
    Set oReadOnlyRads = New colReadOnlyRads
    
    For Each oRad In mCol
        With oRad
        
            bMatched = True
            
            'Compare the dates
            If Filter.RadStartDate <> "" And Filter.RadEndDate <> "" And bMatched = True Then
                If .RadDate >= CDate(Filter.RadStartDate) And .RadDate <= CDate(Filter.RadEndDate) Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            If Filter.RadStartDate <> "" And bMatched = True Then
                If .RadDate >= CDate(Filter.RadStartDate) Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            If Filter.RadEndDate <> "" And bMatched = True Then
                If .RadDate <= CDate(Filter.RadEndDate) Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the Version
            If Filter.Version <> "" And bMatched = True Then
                If oRad.Version = CLng(Filter.Version) Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the RadType
            If Filter.RadType <> "" And bMatched = True Then
                If oRad.RadType = Filter.RadType Then
                    bMatched = True
                ElseIf CompareArrays(oRad.RadType, Filter.RadType) = True Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the Status
            If Filter.Status <> "" And bMatched = True Then
                If oRad.Status = Filter.Status Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the Location
            If Filter.Location <> "" And bMatched = True Then
                'Full match
                If oRad.Location = Filter.Location Then
                    bMatched = True
                'Partial match
                ElseIf InStr(oRad.Location, ", " & Filter.Location & ",") > 0 Then
                    bMatched = True
                'Full search
                ElseIf CompareArrays(oRad.Location, Filter.Location) = True Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the comments
            If Filter.Comments <> "" And bMatched = True Then
                If InStr(1, UCase(oRad.Comments), UCase(Filter.Comments)) > 0 Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the FacilityNCID
            If Filter.FacilityNCID <> "" And bMatched = True Then
                If oRad.FacilityNCID = Filter.FacilityNCID Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the RequestedByNCID
            If Filter.RequestedByNCID <> "" And bMatched = True Then
                If oRad.RequestedByNCID = Filter.RequestedByNCID Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
            'Compare the RequestedByNCID
            If Filter.TakenBy <> "" And bMatched = True Then
                If UCase(oRad.TakenBy) = UCase(Filter.TakenBy) Then
                    bMatched = True
                Else
                    bMatched = False
                End If
            End If
            
        End With
        
        If bMatched = True Then
            oReadOnlyRads.Add oRad
        End If
        
    Next
    
    Set Find = oReadOnlyRads
    Set oReadOnlyRads = Nothing
    
End Function

Friend Function Load(PatientUnitID As Long) As Boolean

'This loads the collection of Rads from the Dental_Imaging_Rads table.

Dim sSQL As String
Dim oRs As Recordset
Dim oRad As clsRad

On Error GoTo ErrHnd
    
    Set moClinicOps = New ClinicOps

    'Reset the data
    Set mCol = New Collection
    mlPatientUnitID = PatientUnitID
    
    'Load the new data
    sSQL = "Select RAD_ID, PATIENT_UNIT_ID, RAD_DATE, VERSION, VERSION_GROUP, REQUESTED_BY_NCID," _
            & " TAKEN_BY, MODIFIED_BY_NCID, LOCATION, FACILITY_NCID, COMMENTS, RADTYPE, STATUS_ID, IMAGE_ID," _
            & " Image_Title, Rad_Order_ID, from Dental_Imaging_Rads Where Patient_Unit_ID = " & PatientUnitID & " and Status_ID <> " & eRadStatus.Deleted & " order by VERSION_GROUP, Rad_ID desc"
                
    Set oRs = goconn.OpenRecordset(sSQL)
    
    Do Until oRs.EOF
        Set oRad = New clsRad
        With oRad
            Set .ClinicOps = moClinicOps
            .RadID = Val(oRs("Rad_ID") & "")
            .RadOrderID = Val(oRs("Rad_Order_ID") & "")
            .PatientUnitID = Val(oRs("Patient_Unit_ID") & "")
            .RadDate = IIf(oRs("Rad_Date") = "12/31/9999", Empty, oRs("Rad_Date") & "")
            .Version = Val(oRs("Version") & "")
            .VersionGroup = Val(oRs("Version_Group") & "")
            .RequestedByNCID = oRs("Requested_By_NCID") & ""
            .TakenBy = oRs("Taken_By") & ""
            .ModifiedByNCID = oRs("Modified_By_NCID") & ""
            .Location = oRs("Location") & ""
            .FacilityNCID = oRs("Facility_NCID") & ""
            .Comments = oRs("Comments") & ""
            .RadType = Val(oRs("RadType") & "")
            .StatusID = Val(oRs("Status_ID") & "")
            .ImageID = oRs("Image_ID") & ""
            .ImageTitle = oRs("Image_Title") & ""
            .DirtyFlag = dfNoAction

            mCol.Add oRad
        End With
        Set oRad = Nothing
        oRs.MoveNext
    Loop
    
    oRs.Close
    Set oRs = Nothing
    Load = True
    
Exit Function

ErrHnd:
    MsgBxARMd "Error in colRads.Load: " & Err.Description
    
End Function

Public Function LookupRadTypeName(RadType As eRadType) As String
        
        LookupRadTypeName = LoadResString(RadType)

End Function

Public Function LookupStatusName(StatusID As eRadStatus) As String

        LookupStatusName = LoadResString(StatusID)

End Function

Public Function RadIconBW() As String
    
    RadIconBW = cRAD_ICON_BW
    
End Function


Public Function RadIconFM() As String

    RadIconFM = cRAD_ICON_FM
    
End Function

Public Function RadIconPA() As String

    RadIconPA = cRAD_ICON_PA
    
End Function


Public Function RadIconPano() As String

    RadIconPano = cRAD_ICON_PANO
    
End Function

Public Function RadIconRadiographs() As String
    
    RadIconRadiographs = cRAD_ICON_RADIOGRAPHS
    
End Function

Public Property Get ImageDir() As String

    ImageDir = GetImagePath
    
End Property

Public Function RadTypeValues() As String
    
    'Returns the possible rad type enums
    RadTypeValues = cRAD_TYPE_VALUES
    
End Function

Public Function StatusIDValues() As String
    
    'Returns the possible rad status id enums
    StatusIDValues = cSTATUS_ID_VALUES
    
End Function

Friend Function Save() As Boolean

'Save all the Rads in the collection

Dim oRad As clsRad

    Save = True
    
    For Each oRad In mCol
        If oRad.Save = False Then
            Save = False
        End If
    Next
    
End Function


Public Function Add(Location As String, RadType As eRadType) As clsRad

    'create a new object
    Dim objNewMember As clsRad
    Set objNewMember = New clsRad
    Dim oRs As Recordset
    
'    'Get the new sequence
'    Set oRs = goconn.OpenRecordset("Select Dental_Imaging_Rads_Seq.NextVal ID from Dual")
'    If oRs.EOF Then
'        MsgBxARMd ("Select Dental_Imaging_Rads_Seq.NextVal ID from Dual did not return a value")
'        Exit Function
'    Else
'        objNewMember.RadID = oRs("ID")
'    End If
'    oRs.Close
'    Set oRs = Nothing
    objNewMember.RadID = goconn.GetNextID(esDENTAL_IMAGING_RADS_SEQ)
    
    With objNewMember
        Set .ClinicOps = moClinicOps
        
        'Set the required properties
        .PatientUnitID = mlPatientUnitID
        
        'set the properties passed into the method
        .Location = Location
        .RadType = RadType
        
        'Set the defaults
        .RequestedByNCID = cwobjLogonEnviron.StoringUserNCID
        .FacilityNCID = cwobjLogonEnviron.StoringFacilityNCID
        .StatusID = eRadStatus.Ordered
        .Version = 0
        .DirtyFlag = dfAdd
    End With
    
    'Add it to the beginning of the collection
    If mCol.Count > 0 Then
        mCol.Add objNewMember, CStr(objNewMember.RadID), 1
    Else
        mCol.Add objNewMember, CStr(objNewMember.RadID)
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function


Public Property Get Item(vntIndexKey As Variant) As clsRad
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Private Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol(vntIndexKey).Status = eRadStatus.Deleted
    
    'mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set moClinicOps = Nothing
    
End Sub

