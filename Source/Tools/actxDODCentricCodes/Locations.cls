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
Attribute VB_Name = "Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Location"
Attribute VB_Ext_KEY = "Member0" ,"Location"
Option Explicit

Private Const MODULE_NAME As String = "Locations"

Private mlFacilityNCID As Long
Private mcLocations As Collection

Friend Function LoadFacilityLocations(ByVal lFacilityNCID As Long) As Boolean

On Error GoTo ErrorHandler

Dim sSql As String
Dim adoRS As ADODB.Recordset
Dim oLocation As Location

   If lFacilityNCID < 0 Then
      Err.Raise 911, , "Invalid Facility NCID"
   End If
   
   'Create facility location collection
   mlFacilityNCID = lFacilityNCID
   
   sSql = "SELECT Location_IEN, Location_Name, Location_Abbreviation " & _
            " FROM Geographic_Location " & _
            " WHERE Facility_NCID = " & lFacilityNCID & _
            " AND Country_IEN is not null " & _
            " ORDER BY Location_Name ASC"
            
   Set adoRS = SharedDAS.OpenRecordset(sSql, adOpenStatic, adLockBatchOptimistic)
   
   If Not adoRS Is Nothing Then
      
      With adoRS
      
         Do While Not .EOF
            'Add each location to collection
            Add .Fields("Location_IEN").value, .Fields("Location_Name").value, .Fields("Location_Abbreviation").value
            .MoveNext
         Loop
         
      End With
      
   End If
      
   LoadFacilityLocations = True
   
ErrorHandler:
   'Clean Up Objects
   Set oLocation = Nothing
   If Not adoRS Is Nothing Then
      If adoRS.State = adStateOpen Then adoRS.Close
      Set adoRS = Nothing
   End If
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "LoadFacilityLocations"
   
End Function

'Adds a new location to the locations collection
Private Function Add(ByVal lIEN As Long, ByVal sName As String, ByVal sAbbrev As String) As Location

Dim oNewLoc As Location

   Set oNewLoc = New Location

   'set the properties passed into the method
   oNewLoc.Name = sName
   oNewLoc.IEN = lIEN
   oNewLoc.Abbreviation = sAbbrev
      
   mcLocations.Add oNewLoc, CStr(oNewLoc.IEN)
   
   'return the object created
   Set Add = oNewLoc
   
   Set oNewLoc = Nothing

End Function

Public Property Get Item(vIndexKey As Variant) As Location
Attribute Item.VB_UserMemId = 0
  Set Item = mcLocations(vIndexKey)
End Property

Public Property Get Count() As Long
   Count = mcLocations.Count
End Property

Friend Sub Remove(vIndexKey As Variant)
   mcLocations.Remove vIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mcLocations.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set mcLocations = New Collection
End Sub

Private Sub Class_Terminate()
   Set mcLocations = Nothing
End Sub

Public Property Get FacilityNCID() As Long
   FacilityNCID = mlFacilityNCID
End Property

