VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Facilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcFacilityLocations As Collection          'Contains a collection of locations collections

'Adds a new facility location collection and returns it
Private Function AddFacility(ByVal lFacilityNCID As Long) As Locations

Dim oLocations As Locations

   Set oLocations = New Locations
   If oLocations.LoadFacilityLocations(lFacilityNCID) Then
      'Load was successfull so add to the collection
      mcFacilityLocations.Add oLocations, CStr(lFacilityNCID)
      Set AddFacility = mcFacilityLocations(CStr(lFacilityNCID))
   Else
      'Error occurred so return an empty collection
      Set AddFacility = New Locations
   End If
   Set oLocations = Nothing
   
End Function

Public Property Get Locations(ByVal lFacilityNCID As Long) As Locations
   If Not IsLoaded(lFacilityNCID) Then
      'Load facility and populate locations
      Set Locations = AddFacility(lFacilityNCID)
   Else
      'Return the previously loaded locations collection
      Set Locations = mcFacilityLocations(CStr(lFacilityNCID))
   End If
End Property

'Indicates if the particular facility has been loaded into the facilities collection
Private Function IsLoaded(ByVal lFacilityNCID As Long) As Boolean
   On Error Resume Next
   IsLoaded = Not mcFacilityLocations(CStr(lFacilityNCID)) Is Nothing
   On Error GoTo 0
End Function

Private Sub Class_Initialize()
   Set mcFacilityLocations = New Collection
End Sub

Private Sub Class_Terminate()
   Set mcFacilityLocations = Nothing
End Sub
