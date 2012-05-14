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
Attribute VB_Name = "cWard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarWardID          As Long
Private mvarSpecialty       As String
Private mvarSpecialtyName   As String
Private mvarLocationncid    As String
Private mvarLocationName    As String
Private mvarBeds            As cBeds
Private mvarStatus          As String
Private mvarFacilityNCID    As String
Private mvarObjectState     As ObjectState_
Public DisplayRow           As Long

Private mvarTotalBeds       As Long
Private mvarOpenBeds        As Long
Private mvarOccupiedBeds    As Long
Private mvarClosedBeds      As Long
Private mvarCapacity        As Long

Public Property Let objectstate(RHV As ObjectState_):       mvarObjectState = RHV:          End Property
Public Property Get objectstate() As ObjectState_:          objectstate = mvarObjectState:  End Property
Public Property Let WardID(ByVal ID As Long):               mvarWardID = ID:                End Property
Public Property Get WardID() As Long:                       WardID = mvarWardID:            End Property
Public Property Let Status(ByVal RHV As String):            mvarStatus = RHV:               End Property
Public Property Get Status() As String:                     Status = mvarStatus:            End Property
Public Property Let FacilityNCID(ByVal RHV As String):      mvarFacilityNCID = RHV:         End Property
Public Property Get FacilityNCID() As String:               FacilityNCID = mvarFacilityNCID:    End Property
Public Property Let LocationNCID(ByVal vData As String):    mvarLocationncid = vData:           End Property
Public Property Get LocationNCID() As String:               LocationNCID = mvarLocationncid:    End Property
Public Property Let LocationName(ByVal vData As String):    mvarLocationName = vData:           End Property
Public Property Get LocationName() As String:               LocationName = mvarLocationName:    End Property
Public Property Get Specialty() As String:                  Specialty = mvarSpecialty:          End Property
Public Property Let SpecialtyName(ByVal vData As String):   mvarSpecialtyName = vData:          End Property
Public Property Get SpecialtyName() As String:              SpecialtyName = mvarSpecialtyName:  End Property
Public Property Set Beds(ByVal vData As cBeds):         Set mvarBeds = vData:                   End Property
Public Property Get Beds() As cBeds:                    Set Beds = mvarBeds:                    End Property
Public Property Get Capacity() As Long:                     Capacity = mvarCapacity:            End Property
Public Property Get OccupiedBeds() As Long:                 OccupiedBeds = mvarOccupiedBeds:    End Property
Public Property Let TotalBeds(ByVal vData As Long):         mvarTotalBeds = vData:              End Property
Public Property Get TotalBeds() As Long:                    TotalBeds = mvarTotalBeds:          End Property
Public Property Let ClosedBeds(ByVal vData As Long):        mvarClosedBeds = vData:             End Property
Public Property Get ClosedBeds() As Long:                   ClosedBeds = mvarClosedBeds:        End Property
Public Property Let OpenBeds(ByVal vData As Long):          mvarOpenBeds = vData:               End Property
Public Property Get OpenBeds() As Long:                     OpenBeds = mvarOpenBeds:            End Property

Public Property Let Specialty(ByVal vData As String)
    mvarSpecialty = vData
    mvarSpecialtyName = gSpecialties.GetSpecialty(vData).SpecName
End Property

Private Sub Class_Initialize()
    Set mvarBeds = New cBeds
End Sub

Private Sub Class_Terminate()
    Set mvarBeds = Nothing
End Sub

Public Function DeserializeFromRecordset(oRS As ADODB.Recordset) As Boolean
    Dim oBed As cBed
    Dim i As Long
    Dim KVP As KeyValuePair
    If oRS.EOF Then Exit Function
    Dim oSpec As cSpecialty
    
    mvarWardID = vbNullString & oRS("Ward_ncid")
    mvarLocationncid = vbNullString & oRS("WARD_LOCATION_NCID")
    mvarLocationName = gcolLocations(CStr(mvarLocationncid)).Value

    mvarSpecialty = vbNullString & oRS("WARD_NAME")
    Set oSpec = gSpecialties.GetSpecialty(mvarSpecialty)
    If Not oSpec Is Nothing Then
        mvarSpecialtyName = gSpecialties.GetSpecialty(mvarSpecialty).SpecName
    ElseIf Not mvarSpecialty = vbNullString Then
        mvarSpecialtyName = "(Not Found)"
    End If
    mvarStatus = vbNullString & oRS("WARD_STATUS_CODE")
    mvarFacilityNCID = vbNullString & oRS("FACILITY_NCID")
End Function

Public Function SerializeToRecordset(oRS As ADODB.Recordset) As Boolean
    
    On Error GoTo ErrHnd
    
    Dim oBed As cBed
    
    If oRS.EOF Then Exit Function
    
    oRS("Ward_ncid") = mvarWardID
    oRS("WARD_LOCATION_NCID") = mvarLocationncid
    oRS("WARD_NAME") = mvarSpecialty
    oRS("WARD_STATUS_CODE") = mvarStatus
    If mvarFacilityNCID = vbNullString Then
        oRS("FACILITY_NCID") = gobjLogon.FacilityNCID
    Else
        oRS("FACILITY_NCID") = mvarFacilityNCID
    End If
    oRS("WARD_BED_NUMBER") = mvarBeds.Count
    
Exit Function
ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cWard.SerializeToRecordset."
    
End Function


Public Function ComputeCapacity() As Boolean
'__ returns the percent full, for this single specilalty
'__ must have specname

    Dim TotalBeds As Long
    Dim OpenBeds As Long
    Dim Ward As cWard
    Dim Bed As cBed
    Dim OccBeds As Long
    Dim pat As cInpatient
    On Error GoTo ErrHandler
    If gWards Is Nothing Then Exit Function  '__ happens on inital load
    
'__ count the total number of beds in the specialty
'   and collect up the wards, too
    
    mvarTotalBeds = mvarBeds.Count
    mvarOccupiedBeds = 0
    mvarClosedBeds = 0
    mvarOpenBeds = 0
    
    If mvarStatus = "INACTIVE" Then
        Exit Function
    End If
    '__ some may be closed, some may be inactive.  look at each bed
   
    For Each Bed In mvarBeds
        If Bed.Status = "CLOSED" Then
            mvarClosedBeds = mvarClosedBeds + 1
        End If
        If Not Bed.Patient Is Nothing Then
            If Not Val(Bed.Patient.Unit_Number) = 0 Then
                mvarOccupiedBeds = mvarOccupiedBeds + 1
            End If
        End If
    Next Bed
    
'__ computer the percent capacity of the specialty
    mvarOpenBeds = mvarBeds.Count - mvarClosedBeds
    
    If mvarOpenBeds > 0 Then
        mvarCapacity = Int((mvarOccupiedBeds / mvarOpenBeds) * 100)
    ElseIf mvarOccupiedBeds > 0 Then
        mvarCapacity = 100
    Else
        mvarCapacity = 0
    End If
    
    ComputeCapacity = True
    Exit Function
    
ErrHandler:
    ComputeCapacity = False
End Function

