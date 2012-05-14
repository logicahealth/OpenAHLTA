VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Application reference item in tree structure.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "Clinic"
Private Const OBJECT_VERSION As String = "1.0"

Private mstrFacilityNCID           As String    'Clinic Facility NCID
Private mstrNCID                   As String    'Clinic NCID
Private mstrName                   As String    'Clinic Name
Private mstrIEN                    As String    'Clinic IEN
Private mstrMeprsCode              As String    'Clinic Meprs Code
Private mstrMeprsDescription       As String    'Clinic Meprs Description
Private mstrMeprsIEN               As String    'Clinic Meprs IEN
Private mstrLocationType           As String    'Clinic Location Type
Private mbWalkInAllowed            As Boolean   'Clinic Walk In Allowed
Private mbSickCallAllowed          As Boolean   'Clinic Sick Call Allowed
Private mstrStreetAddress          As String    'Clinic Street Address
Private mstrBuildingNumber         As String    'Clinic Building Number
Private mstrCity                   As String    'Clinic City
Private mstrState                  As String    'Clinic State
Private mstrZipCode                As String    'Clinic ZIP Code
Private mstrDmisIdCode             As String    'Clinic DMIS Id Code
Private mstrDefPharmacyBDM         As String    'Clinic Defined Pharmacy BDM
Private mstrDivisionStateIEN       As String    'Clinic Division State IEN

'''Workload properties
Private mstrDefault_Workload_Type  As String    'Clinic Workload Default Type
Private mcolWorkloads              As Workloads 'Clinic Workload Collection

Private Function iObjectPersist_Copy() As Object
Dim Item As Clinic
Dim objPersist As iObjectPersist
  On Error GoTo ErrHandler

  Set Item = New Clinic
  
  With Item
    .FacilityNCID = mstrFacilityNCID
    .NCID = mstrNCID
    .Name = mstrName
    .IEN = mstrIEN
    .MeprsCode = mstrMeprsCode
    .MeprsDescription = mstrMeprsDescription
    .MeprsIEN = mstrMeprsIEN
    .LocationType = mstrLocationType
    .WalkInAllowed = mbWalkInAllowed
    .SickCallAllowed = mbSickCallAllowed
    .StreetAddress = mstrStreetAddress
    .BuildingNumber = mstrBuildingNumber
    .City = mstrCity
    .State = mstrState
    .ZipCode = mstrZipCode
    .DmisIdCode = mstrDmisIdCode
    .DefPharmacyBDM = mstrDefPharmacyBDM
    .Default_Workload_Type = mstrDefault_Workload_Type
    .DivisionStateIEN = mstrDivisionStateIEN
    
    Set objPersist = New Workloads
    Set .ClinicWorkloads = objPersist.Copy
  End With
  
  Set iObjectPersist_Copy = Item
  Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "DTOsClinic.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  On Error GoTo ErrHandler
  
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    .WriteProperty "FacilityNCID", mstrFacilityNCID
    .WriteProperty "NCID", mstrNCID
    .WriteProperty "Name", mstrName
    .WriteProperty "IEN", mstrIEN
    .WriteProperty "MeprsCode", mstrMeprsCode
    .WriteProperty "MeprsDescription", mstrMeprsDescription
    .WriteProperty "MeprsIEN", mstrMeprsIEN
    .WriteProperty "LocationType", mstrLocationType
    .WriteProperty "WalkInAllowed", CInt(mbWalkInAllowed)
    .WriteProperty "SickCallAllowed", CInt(mbSickCallAllowed)
    .WriteProperty "StreetAddress", mstrStreetAddress
    .WriteProperty "BuildingNumber", mstrBuildingNumber
    .WriteProperty "City", mstrCity
    .WriteProperty "State", mstrState
    .WriteProperty "ZipCode", mstrZipCode
    .WriteProperty "DmisIdCode", mstrDmisIdCode
    .WriteProperty "DefPharmacyBDM", mstrDefPharmacyBDM
    .WriteProperty "Default_Workload_Type", mstrDefault_Workload_Type
    .WriteProperty "DivisionStateIEN", mstrDivisionStateIEN
    
    .WriteObject "Workloads", mcolWorkloads
  
  End With
  Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsClinic.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mstrIEN = .ReadProperty("IEN")
    mstrFacilityNCID = .ReadProperty("FacilityNCID")
    mstrNCID = .ReadProperty("NCID")
    mstrName = .ReadProperty("Name")
    mstrIEN = .ReadProperty("IEN")
    mstrMeprsCode = .ReadProperty("MeprsCode")
    mstrMeprsDescription = .ReadProperty("MeprsDescription")
    mstrMeprsIEN = .ReadProperty("MeprsIEN")
    mstrLocationType = .ReadProperty("LocationType")
    mbWalkInAllowed = CBool(.ReadProperty("WalkInAllowed"))
    mbSickCallAllowed = CBool(.ReadProperty("SickCallAllowed"))
    mstrStreetAddress = .ReadProperty("StreetAddress")
    mstrBuildingNumber = .ReadProperty("BuildingNumber")
    mstrCity = .ReadProperty("City")
    mstrState = .ReadProperty("State")
    mstrZipCode = .ReadProperty("ZipCode")
    mstrDmisIdCode = .ReadProperty("DmisIdCode")
    mstrDefPharmacyBDM = .ReadProperty("DefPharmacyBDM")
    mstrDefault_Workload_Type = .ReadProperty("Default_Workload_Type")
    mstrDivisionStateIEN = .ReadProperty("DivisionStateIEN")
    
    .ReadObject "Workloads", ClinicWorkloads
  End With

  iObjectPersist_Restore = Not Failed
  
  Exit Function
  
ErrHandler:
  Failed = True
  Resume Next

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Property Get FacilityNCID() As String
  FacilityNCID = mstrFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal NewValue As String)
  mstrFacilityNCID = NewValue
End Property

Public Property Get NCID() As String
  NCID = mstrNCID
End Property
Public Property Let NCID(ByVal NewValue As String)
  mstrNCID = NewValue
End Property

Public Property Get Name() As String
  Name = mstrName
End Property
Public Property Let Name(ByVal NewValue As String)
  mstrName = NewValue
End Property

Public Property Get IEN() As String
  IEN = mstrIEN
End Property
Public Property Let IEN(ByVal NewValue As String)
  mstrIEN = NewValue
End Property

Public Property Get MeprsCode() As String
  MeprsCode = mstrMeprsCode
End Property
Public Property Let MeprsCode(ByVal NewValue As String)
  mstrMeprsCode = NewValue
End Property

Public Property Get MeprsDescription() As String
  MeprsDescription = mstrMeprsDescription
End Property
Public Property Let MeprsDescription(ByVal NewValue As String)
  mstrMeprsDescription = NewValue
End Property

Public Property Get MeprsIEN() As String
  MeprsIEN = mstrMeprsIEN
End Property
Public Property Let MeprsIEN(ByVal NewValue As String)
  mstrMeprsIEN = NewValue
End Property

Public Property Get LocationType() As String
  LocationType = mstrLocationType
End Property
Public Property Let LocationType(ByVal NewValue As String)
  mstrLocationType = NewValue
End Property

Public Property Get WalkInAllowed() As Boolean
  WalkInAllowed = mbWalkInAllowed
End Property
Public Property Let WalkInAllowed(ByVal NewValue As Boolean)
  mbWalkInAllowed = NewValue
End Property

Public Property Get SickCallAllowed() As Boolean
  SickCallAllowed = mbSickCallAllowed
End Property
Public Property Let SickCallAllowed(ByVal NewValue As Boolean)
  mbSickCallAllowed = NewValue
End Property

Public Property Get StreetAddress() As String
  StreetAddress = mstrStreetAddress
End Property
Public Property Let StreetAddress(ByVal NewValue As String)
  mstrStreetAddress = NewValue
End Property

Public Property Get BuildingNumber() As String
  BuildingNumber = mstrBuildingNumber
End Property
Public Property Let BuildingNumber(ByVal NewValue As String)
  mstrBuildingNumber = NewValue
End Property

Public Property Get City() As String
  City = mstrCity
End Property
Public Property Let City(ByVal NewValue As String)
  mstrCity = NewValue
End Property

Public Property Get State() As String
  State = mstrState
End Property
Public Property Let State(ByVal NewValue As String)
  mstrState = NewValue
End Property

Public Property Get ZipCode() As String
  ZipCode = mstrZipCode
End Property
Public Property Let ZipCode(ByVal NewValue As String)
  mstrZipCode = NewValue
End Property

Public Property Get DmisIdCode() As String
  DmisIdCode = mstrDmisIdCode
End Property
Public Property Let DmisIdCode(ByVal NewValue As String)
  mstrDmisIdCode = NewValue
End Property

Public Property Get DefPharmacyBDM() As String
  DefPharmacyBDM = mstrDefPharmacyBDM
End Property
Public Property Let DefPharmacyBDM(ByVal NewValue As String)
  mstrDefPharmacyBDM = NewValue
End Property

Public Property Get Default_Workload_Type() As String
  Default_Workload_Type = mstrDefault_Workload_Type
End Property
Public Property Let Default_Workload_Type(ByVal NewValue As String)
  mstrDefault_Workload_Type = NewValue
End Property

Public Property Get DivisionStateIEN() As String
  DivisionStateIEN = mstrDivisionStateIEN
End Property
Public Property Let DivisionStateIEN(ByVal NewValue As String)
  mstrDivisionStateIEN = NewValue
End Property

Public Property Get ClinicWorkloads() As Workloads
  If mcolWorkloads Is Nothing Then
    Set mcolWorkloads = New Workloads
  End If
  
  Set ClinicWorkloads = mcolWorkloads
End Property
Public Property Set ClinicWorkloads(ByVal NewValue As Workloads)
  Set mcolWorkloads = NewValue
End Property


