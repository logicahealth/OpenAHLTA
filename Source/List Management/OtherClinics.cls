VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOwners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ListOwner"
Attribute VB_Ext_KEY = "Member0" ,"ListOwner"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Property Get Item(vntIndexKey As Variant) As ListOwner
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
End Sub

Friend Sub Init(ByVal AllowRead As Boolean, ByVal AllowWrite As Boolean, ByVal AllowDelete As Boolean, Optional ByVal ProviderNCID As String = "")


    Dim objMyClinics As DTOs.Clinics
    Dim objClinicOps As CHCSII_ClinicClient.ClinicOpsEx
    Dim objClinic As DTOs.Clinic
    Dim objListOwner As ListOwner
    
    Set mCol = New Collection
    
    Set objClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
    
    'SCR 130196
    If ProviderNCID = "" Then
        ProviderNCID = cwobjLogonEnviron.UserNCID
    End If
    
    Set objMyClinics = objClinicOps.GetUserClinics(cwobjLogonEnviron.FacilityNCID, ProviderNCID)
    
    For Each objClinic In objMyClinics
        'SCR-44069 RBELL 10/14/03 Added long conversion as some of the values from the logon object had a leading zero and were failing the comparison
        If CLng(objClinic.NCID) <> CLng(gobjLogon.NursingDivisionNCID) Then
            Set objListOwner = New ListOwner
            objListOwner.Init otClinic, objClinic.NCID, gobjLogon.FacilityNCID
            mCol.Add objListOwner, CStr(objClinic.NCID)
        End If
    Next

End Sub
Friend Sub Add(ByRef objListOwner As ListOwner, ByVal OwnerNCID As Long)

    mCol.Add objListOwner, CStr(OwnerNCID)

End Sub
