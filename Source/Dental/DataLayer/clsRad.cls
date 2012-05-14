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
Attribute VB_Name = "clsRad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This is the basic object for a Radiograph typically referenced from colRads

Private mlRadID As Long
Private mlRadOrderID As Long
Private mlPatientUnitID As Long
Private mdRadDate As Date
Private mlVersion As Long
Private mlVersionGroup As Long

Private msRequestedBy As String
Private mlRequestedByNCID As Long
'Used to cache the RequestedBy name
Private mlRequestedByNCIDPrev As Long

Private msTakenBy As String

Private msModifiedBy As String
Private mlModifiedByNCID As Long
'Used to cache the ModifiedBy name
Private mlModifiedByNCIDPrev As Long

Private msLocation As String 'Location of the image content in the mouth. Teeth:(1, 2, 3, etc.)
Private msFacility As String
Private mlFacilityNCID As Long
Private mlFacilityNCIDPrev As Long

Private msComments As String
Private meRadType As eRadType
Private msRadIcon As String
Private meStatusID As eRadStatus
Private msRadTypeName As String
Private msStatus As String
Private meDirtyFlag As DirtyFlagEnum
Private mlImageID As Long

'For now, these are the file titles (file name only)
Private msImageTitle As String

Private moClinicOps As ClinicOps


Public Property Get AnnotationTitle() As String

'This function gets the annotation filename from the image filename

Dim iPos As Integer

    If msImageTitle <> "" Then
        iPos = InStrRev(msImageTitle, ".")
        If iPos > 0 Then
            AnnotationTitle = Mid(msImageTitle, 1, iPos - 1) & ".art"
        Else
            AnnotationTitle = msImageTitle & ".art"
        End If
    End If
    
End Property

Public Property Get FacilityNCID() As Long

    FacilityNCID = mlFacilityNCID
    
End Property

Public Property Let FacilityNCID(vData As Long)

    mlFacilityNCID = vData
    
    'Get the name of the facility
    If mlFacilityNCIDPrev <> mlFacilityNCID Then
        'Add logic to translate a facility NCID to a facility name
    End If
    
    SetDirtyFlag
    
End Property

Friend Property Let PatientUnitID(vData As Long)

    mlPatientUnitID = vData
    SetDirtyFlag
    
End Property

Public Property Get PatientUnitID() As Long
    
    PatientUnitID = mlPatientUnitID
    
End Property


Public Property Get RadIcon() As String
    
    RadIcon = msRadIcon
    
End Property

Public Property Let RadOrderID(vData As Long)
    
    mlRadOrderID = vData
    
End Property

Public Property Get RadOrderID() As Long
    
    RadOrderID = mlRadOrderID
    
End Property

Private Sub SetDirtyFlag()
    
    'This is called when a user sets a property
    'If the current flag is dfAdd, then don't change it until Save is called
    'otherwise, set it to sfUpdate
    
    If meDirtyFlag = dfNoAction Then
        meDirtyFlag = dfUpdate
    End If
    
End Sub


Friend Property Let RadID(ByVal vData As Long)
    mlRadID = vData
    SetDirtyFlag
End Property

Public Property Get RadID() As Long
    RadID = mlRadID
End Property

Public Property Let ImageTitle(ByVal vData As String)
    msImageTitle = vData
    SetDirtyFlag
End Property

Public Property Get ImageTitle() As String
    
    If msImageTitle = "" Then
        'If there is no image title, then create one.
        msImageTitle = Me.PatientUnitID & "-" & Format(Me.RadID, "000000") & ".dcm"
        SetDirtyFlag
    End If
    
    ImageTitle = msImageTitle
    
End Property

Public Property Let ImageID(ByVal vData As Long)
    mlImageID = vData
    SetDirtyFlag
End Property

Public Property Get ImageID() As Long
    ImageID = mlImageID
End Property

Friend Property Let DirtyFlag(ByVal vData As DirtyFlagEnum)
    meDirtyFlag = vData
End Property

Friend Property Get DirtyFlag() As DirtyFlagEnum
    DirtyFlag = meDirtyFlag
End Property

Friend Function Save() As Boolean
'Save the object to the database

Dim sSQL As String
Dim oRs As Recordset

On Error GoTo ErrHnd

    Select Case meDirtyFlag
    
        Case DirtyFlagEnum.dfNoAction
            Save = True
            Exit Function
            
        Case DirtyFlagEnum.dfAdd
        
'            'Get the sequence
'            Set oRs = goconn.OpenRecordset("Select Dental_Imaging_Rads_Seq.NextVal ID from Dual")
'            If oRs.EOF Then
'                MsgBxARMd ("Select Dental_Imaging_Rads_Seq.NextVal ID from Dual did not return a value")
'                Save = False
'                Exit Function
'            Else
'                mlRadID = oRs("ID")
'            End If
'            oRs.Close
'            Set oRs = Nothing
            mlRadID = goconn.GetNextID(esDENTAL_IMAGING_RADS_SEQ)
            
            'Create the insert statement
            sSQL = "Insert into Dental_Imaging_Rads (RAD_ID, Rad_Order_ID, PATIENT_UNIT_ID, RAD_DATE, VERSION, VERSION_GROUP, REQUESTED_BY_NCID," _
                & " TAKEN_BY, MODIFIED_BY_NCID, LOCATION, FACILITY_NCID, COMMENTS, RADTYPE, STATUS_ID, IMAGE_ID," _
                & " Image_Title) Values (" & mlRadID & ", " & mlRadOrderID & ", " & mlPatientUnitID & ", " & goChcsiiConn.FormatDateForDBInsert(mdRadDate) & ", " & mlVersion & ", " & mlVersionGroup & "," & mlRequestedByNCID & "" _
                & ", '" & msTakenBy & "', " & mlModifiedByNCID & " , '" & msLocation & "', " & mlFacilityNCID & ", '" & msComments & "', " & meRadType & ", " & meStatusID & ", " & mlImageID _
                & ", '" & Me.ImageTitle & "')"
            
        Case DirtyFlagEnum.dfdelete
            'Create the delete statement
            If mlRadID > 0 Then
                sSQL = "Delete from Dental_Imaging_Rads where Rad_ID = " & mlRadID
            End If
            
        Case DirtyFlagEnum.dfUpdate
            'Create the update statement
            sSQL = "Update Dental_Imaging_Rads Set " _
                & " Rad_Order_ID = " & mlRadOrderID & "," _
                & " RAD_DATE = " & goChcsiiConn.FormatDateForDBInsert(mdRadDate) & "," _
                & " VERSION = " & mlVersion & "," _
                & " VERSION_GROUP = " & mlVersionGroup & "," _
                & " REQUESTED_BY_NCID = '" & mlRequestedByNCID & "'," _
                & " TAKEN_BY = '" & msTakenBy & "'," _
                & " MODIFIED_BY_NCID = '" & mlModifiedByNCID & "'," _
                & " LOCATION = '" & msLocation & "'," _
                & " FACILITY_NCID = '" & mlFacilityNCID & "'," _
                & " COMMENTS = '" & msComments & "'," _
                & " RADTYPE = " & meRadType & "," _
                & " STATUS_ID = " & meStatusID & "," _
                & " IMAGE_ID = " & mlImageID & "," _
                & " Image_Title = '" & Me.ImageTitle & "' where Rad_ID = " & mlRadID
            
    End Select
            
    If sSQL <> "" Then
        goconn.ExecuteSQL sSQL
    End If
    
    meDirtyFlag = dfNoAction
    
    Save = True
    
Exit Function

ErrHnd:
    MsgBxARMd "Error in clsRad.Save: " & Err.Description
    
End Function

Private Property Let Status(ByVal vData As String)
    msStatus = vData
    SetDirtyFlag
End Property

Public Property Get Status() As String
    Status = msStatus
End Property

Private Property Let RadTypeName(ByVal vData As String)
    msRadTypeName = vData
    SetDirtyFlag
End Property

Public Property Get RadTypeName() As String
    RadTypeName = msRadTypeName
End Property

Public Property Let StatusID(ByVal vData As eRadStatus)

On Error GoTo ErrHnd

    msStatus = LoadResString(vData)
    meStatusID = vData
    
    If meStatusID = eRadStatus.Deleted Then
        meDirtyFlag = dfdelete
    End If
    
    SetDirtyFlag
    
Exit Property

ErrHnd:
    MsgBxARMd "Could not set the clsRad.StatusID property to: " & vData

End Property

Public Property Get StatusID() As eRadStatus
    Status = meStatusID
End Property

Public Property Let RadType(ByVal vData As eRadType)
    On Error GoTo ErrHnd
    
    msRadTypeName = LoadResString(vData)
    meRadType = vData
    
    'Set the radicon
    Select Case meRadType
        Case eRadType.BiteWing
            msRadIcon = cRAD_ICON_BW
        Case eRadType.FMSeries
            msRadIcon = cRAD_ICON_FM
        Case eRadType.PA
            msRadIcon = cRAD_ICON_PA
        Case eRadType.Pano
            msRadIcon = cRAD_ICON_PANO
        Case eRadType.Other
            msRadIcon = cRAD_ICON_PANO

    End Select
    
    SetDirtyFlag
    
Exit Property

ErrHnd:
    MsgBxARMd "Could not set the clsRad.RadType property to: " & vData
    
End Property

Public Property Get RadType() As eRadType
    RadType = meRadType
End Property

Public Property Let Comments(ByVal vData As String)
    msComments = vData
    SetDirtyFlag
End Property

Public Property Get Comments() As String
    Comments = msComments
End Property

Public Property Let Facility(ByVal vData As String)
    msFacility = vData
    SetDirtyFlag
End Property

Public Property Get Facility() As String
    Facility = msFacility
End Property

Public Property Let Location(ByVal vData As String)
    msLocation = vData
    SetDirtyFlag
End Property

Public Property Get Location() As String
    Location = msLocation
End Property

Public Property Let ModifiedByNCID(ByVal vData As Long)
    
    mlModifiedByNCID = vData

    'Get the string, Logic to cache the results
    If mlModifiedByNCIDPrev <> mlModifiedByNCID Then
        msModifiedBy = moClinicOps.GetProviderName(cwobjLogonEnviron.FacilityNCID, mlModifiedByNCID)
        mlModifiedByNCIDPrev = mlModifiedByNCID
    End If

    SetDirtyFlag
    
End Property

Public Property Get ModifiedByNCID() As Long
    ModifiedByNCID = mlModifiedByNCID
End Property

Private Property Let ModifiedBy(ByVal vData As String)
'    msModifiedBy = vData
End Property

Public Property Get ModifiedBy() As String
    
    ModifiedBy = msModifiedBy

End Property

Public Property Let TakenBy(ByVal vData As String)
    msTakenBy = vData
    SetDirtyFlag
End Property

Public Property Get TakenBy() As String
    'This needs to be determined by the NCID
    TakenBy = msTakenBy
End Property

Public Property Let RequestedByNCID(ByVal vData As Long)
    
    mlRequestedByNCID = vData

    'Get the string, Logic to cache the results
    If mlRequestedByNCIDPrev <> mlRequestedByNCID Then
        If Not cwobjLogonEnviron Is Nothing Then
            msRequestedBy = moClinicOps.GetProviderName(cwobjLogonEnviron.FacilityNCID, mlRequestedByNCID)
            mlRequestedByNCIDPrev = mlRequestedByNCID
        End If
    End If

    SetDirtyFlag
    
End Property

Public Property Get RequestedByNCID() As Long
    RequestedByNCID = mlRequestedByNCID
End Property

Private Property Let RequestedBy(ByVal vData As String)
'    msRequestedBy = vData
End Property

Public Property Get RequestedBy() As String
    'This needs to be determined by the NCID
    
    RequestedBy = msRequestedBy
    
End Property

Public Property Get ThumbnailTitle() As String

'This function gets the thumnail filename from the image filename

Dim iPos As Integer

    If msImageTitle <> "" Then
        iPos = InStrRev(msImageTitle, ".")
        If iPos > 0 Then
            ThumbnailTitle = Mid(msImageTitle, 1, iPos - 1) & ".thm"
        Else
            ThumbnailTitle = msImageTitle & ".thm"
        End If
    End If

End Property

Public Property Get ThumnailPic() As IPictureDisp
    
'This gets a ole error

'Return a picture object of the thumbnail.
'Need to cache this later

'On Error GoTo ErrHnd
'
'    If msImageTitle <> "" Then
'        If Dir(GetImageTitle & "\" & Me.ThumbnailTitle) <> "" Then
'            'The thumbnail exists so load it
'            Set ThumnailPic = LoadPicture(GetImageTitle & "\" & Me.ThumbnailTitle)
'        Else
'            'The thumbnail does not exist so load the real image as the thumnail
'            Set ThumnailPic = LoadPicture(GetImageTitle & "\" & msImageTitle)
'        End If
'    End If
'
'Exit Property
'
'ErrHnd:
'    MsgBox "Error in property get clsRad.ThumbnailPic: " & Err.Description & " Error number: " & Err.Number
'    Resume Next
    
End Property

Public Property Let Version(ByVal vData As Long)
    mlVersion = vData
    SetDirtyFlag
End Property

Public Property Get ImageDir() As String

    ImageDir = GetImagePath
    
End Property

Public Property Get Version() As Long
    Version = mlVersion
End Property

Public Property Let RadDate(ByVal vData As Date)
    mdRadDate = vData
    SetDirtyFlag
End Property

Public Property Get RadDate() As Date
    RadDate = mdRadDate
End Property



Public Property Let VersionGroup(vData As Long)

    mlVersionGroup = vData
    SetDirtyFlag
    
End Property

Public Property Get VersionGroup() As Long
    
    VersionGroup = mlVersionGroup
    
End Property

Public Property Set ClinicOps(vData As ClinicOps)

    Set moClinicOps = vData
    
End Property

Private Sub Class_Terminate()

    Set moClinicOps = Nothing
    
End Sub


