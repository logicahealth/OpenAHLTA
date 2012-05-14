VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "thr_MmiDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moMember        As IMember
Private moMemberG       As CHCSII_MMI.thr_Member     'Gems adaptor interface

Private mbLock          As Boolean

Private moGemsMMIops    As GEMS_MMIOps

'***********************************************************
' Collection class for Address, Demographic, Facility,
' GenericData, GenericID, Name, Phone, and User.
'
'***********************************************************

Implements IMmiDoc

Private Sub Class_Initialize()

    Set moGemsMMIops = New GEMS_MMIOps
    
End Sub

Private Function IMmiDoc_Add() As Long
    Debug.Assert 0
End Function

Private Sub IMmiDoc_ApplicationName(ByVal lpszName As String)
    Debug.Assert 0
End Sub

Private Sub IMmiDoc_ApplicationVersion(ByVal lpszVersion As String)
    Debug.Assert 0
End Sub

Private Function IMmiDoc_Bobbit() As Boolean
    Debug.Assert 0
End Function

Private Property Get IMmiDoc_BOF() As Boolean
    Debug.Assert 0
End Property
'use lpszNewValue = "Sponsor" so the MmiDoc.Get knows about it.
'
'                                       lpszNewValue is usually "Patient"
Private Sub IMmiDoc_CreateMember(ByVal lpszNewValue As String)
    
    Set moMember = New thr_Member
    Set moMemberG = moMember
        
    Set moMemberG.GemsMMI = moGemsMMIops.CurrentPatient
    
    If UCase$(lpszNewValue) = "SPONSOR" Then
        moMemberG.bSponsor = True
    End If

End Sub

Private Function IMmiDoc_DoSearch(ByVal SearchBy As Object) As Boolean
'    Set moMember = New thr_Member
'    Set moMemberG = moMember
        
    If moGemsMMIops.DoSearch(SRCH_ID_None, SRCH_SSN, "", Right(SearchBy.Member.SSN, 4), "", "", "", "", SearchBy.Member.SSN, "", "", "", "", SRCH_Generic_None, "").Count Then
        IMmiDoc_DoSearch = True
        
        Set moMemberG.GemsMMI = moGemsMMIops.CurrentPatient
    Else
        IMmiDoc_DoSearch = False
    End If
'    Debug.Assert 0
End Function

Private Function IMmiDoc_EndSearch() As Boolean
    Debug.Assert 0
End Function

Private Property Get IMmiDoc_EOF() As Boolean
    Debug.Assert 0
End Property
'************************************************************
' set either UnitNumber or one of the search criteria objects
' before calling this method. DoSearch takes a criteria object
' as a parameter.
'
' Criteria objects:
'    AllergyDataAccess
'    DocumentDataAccess
'    EventActionDataAccess
'    GeneralDataAccess
'    LaboratoryDataAccess
'    ObservationDataAccess
'    OrderDataAccess
'    PharmacyDataAccess
'************************************************************
Private Function IMmiDoc_Get(ByVal bLock As Long) As Boolean
    Dim moGemsMMI As GEMS_MMI

    'if bLock then write mode
    mbLock = bLock
        
    Set moGemsMMIops = New GEMS_MMIOps
    
    'false to set Current Patient, True to get a Sponsor
    If moMemberG.bSponsor Then
        Set moGemsMMI = moGemsMMIops.GetPatient(moMember.UnitNumber, True)
    Else
        Set moGemsMMI = moGemsMMIops.GetPatient(moMember.UnitNumber, False)
    End If

    If moGemsMMI Is Nothing Then
        IMmiDoc_Get = False
    Else
        IMmiDoc_Get = True
        Set moMemberG.GemsMMI = moGemsMMI
    End If

End Function

Private Function IMmiDoc_IsBOF() As Boolean
    Debug.Assert 0
End Function

Private Function IMmiDoc_IsEOF() As Boolean
    Debug.Assert 0
End Function

Private Property Get IMmiDoc_Member() As Object

    ' from database or some persistent store

    Set IMmiDoc_Member = moMember

End Property

Private Sub IMmiDoc_Quit()
        Debug.Assert 0
End Sub

Private Function IMmiDoc_ReleaseLock() As Boolean

    mbLock = False

End Function

Private Function IMmiDoc_Remove() As Boolean
    Debug.Assert 0
End Function

Private Sub IMmiDoc_RemoveAll()
    Debug.Assert 0
End Sub

Private Property Get IMmiDoc_User() As Object
    Debug.Assert 0
End Property

Private Function IMmiDoc_WriteUpdate() As Boolean

    IMmiDoc_WriteUpdate = moGemsMMIops.PutPatient(moMemberG.GemsMMI)

End Function

