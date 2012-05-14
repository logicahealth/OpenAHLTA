VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatientSearchParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Implements iParameters
Implements iObjectPersist

Private Const cUNITNUM As String = "UnitNumber"
Private Const cMEDRECNUMBER As String = "MedicalRecordNumber"
Private Const cLNAME As String = "LastName"
Private Const cFNAME As String = "FirstName"
Private Const cDOB As String = "DateOfBirth"
Private Const cSSN As String = "SSN"
Private Const cSPONSSN As String = "SponsorSSN"
Private Const cSEX As String = "Sex"
Private Const cUIC As String = "UIC"
Private Const cFMP As String = "FMP"
Private Const cFACNCID As String = "FacilityNCID"
Private Const cINCSPONSOR As String = "IncludeSponsor"
Private Const cINCBASIC As String = "IncludeBasicOnly"
Private Const cQUKSEARCH As String = "QuickSearch"
Private Const cLASTRECID As String = "LastRecordID"
Private Const cMAXNUMRECORDS As String = "MaximumNumberOfRecords"
Private Const cCHCSSRH As String = "SearchCHCS"
Private Const cGETALLPATS As String = "GetAllPatients"
Private Const cShowVIP As String = "ShowVIP"
Private Const cShowADSFlag As String = "SearchADS"
Private Const cLoadUnverified As String = "LoadUnverified" '--- SCR 84911  Sherry Wang 1/6/2006

Private mParams As iParameters                      'ParametersGeneric class used to store and persist the NCIDs

Private Sub Class_Initialize()
  Set mParams = New ParametersGeneric
End Sub

Private Function iObjectPersist_Copy() As Object
Dim Persist As iObjectPersist

  Set Persist = mParams
  Set iObjectPersist_Copy = Persist.Copy
  
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
Dim Persist As iObjectPersist

  Set Persist = mParams
  Persist.Persist Writer
End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
Dim Persist As iObjectPersist

  Set Persist = mParams
  iObjectPersist_Restore = Persist.Restore(Reader)
  
End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
Dim Persist As iObjectPersist

  Set Persist = mParams
  iObjectPersist_UniqueIdentifier = Persist.UniqueIdentifier
  
End Property

Private Property Get iObjectPersist_Version() As String
Dim Persist As iObjectPersist

  Set Persist = mParams
  iObjectPersist_Version = Persist.Version
  
End Property

Private Function iParameters_Add(ByVal Name As String, ByVal Value As Variant) As DL_Support.Parameter
  Set iParameters_Add = mParams.Add(Name, Value)
End Function

Private Function iParameters_Count() As Long
  iParameters_Count = mParams.Count
End Function

Private Function iParameters_Exists(ByVal Name As String) As Boolean
  iParameters_Exists = mParams.Exists(Name)
End Function

Private Function iParameters_Item(ByVal Key As Variant) As DL_Support.Parameter
  Set iParameters_Item = mParams.Item(Key)
End Function

Private Sub iParameters_Remove(ByVal Key As Variant)
  mParams.Remove Key
End Sub

Private Function GetValue(ByVal Name As String) As String
  If mParams.Exists(Name) Then
    GetValue = mParams.Item(Name).Value
  Else
    GetValue = vbNullString
  End If
End Function

Private Sub SetValue(ByVal Name As String, ByVal NewValue As String)
  If mParams.Exists(Name) Then
    mParams.Item(Name).Value = NewValue
  Else
    mParams.Add Name, NewValue
  End If
End Sub




Public Property Get UnitNumber() As String
    UnitNumber = GetValue(cUNITNUM)
End Property
Public Property Let UnitNumber(ByVal NewValue As String)
    SetValue cUNITNUM, NewValue
End Property
Public Property Get MedicalRecordNumber() As String
    MedicalRecordNumber = GetValue(cMEDRECNUMBER)
End Property
Public Property Let MedicalRecordNumber(ByVal NewValue As String)
    SetValue cMEDRECNUMBER, NewValue
End Property
Public Property Get LastName() As String
    LastName = GetValue(cLNAME)
End Property
Public Property Let LastName(ByVal NewValue As String)
    SetValue cLNAME, NewValue
End Property
Public Property Get FirstName() As String
    FirstName = GetValue(cFNAME)
End Property
Public Property Let FirstName(ByVal NewValue As String)
    SetValue cFNAME, NewValue
End Property
Public Property Get DOB() As String
    DOB = GetValue(cDOB)
End Property
Public Property Let DOB(ByVal NewValue As String)
    SetValue cDOB, NewValue
End Property
Public Property Get SSN() As String
    SSN = GetValue(cSSN)
End Property
Public Property Let SSN(ByVal NewValue As String)
    SetValue cSSN, NewValue
End Property
Public Property Get SponsorSSN() As String
    SponsorSSN = GetValue(cSPONSSN)
End Property
Public Property Let SponsorSSN(ByVal NewValue As String)
    SetValue cSPONSSN, NewValue
End Property
Public Property Get SEX() As String
    SEX = GetValue(cSEX)
End Property
Public Property Let SEX(ByVal NewValue As String)
    SetValue cSEX, NewValue
End Property
Public Property Get UIC() As String
    UIC = GetValue(cUIC)
End Property
Public Property Let UIC(ByVal NewValue As String)
    SetValue cUIC, NewValue
End Property
Public Property Get FMP() As String
    FMP = GetValue(cFMP)
End Property
Public Property Let FMP(ByVal NewValue As String)
    SetValue cFMP, NewValue
End Property
Public Property Get FacilityNCID() As String
    FacilityNCID = GetValue(cFACNCID)
End Property
Public Property Let FacilityNCID(ByVal NewValue As String)
    SetValue cFACNCID, NewValue
End Property
Public Property Get IncludeSponsor() As Boolean
    If GetValue(cINCSPONSOR) = vbNullString Then
        IncludeSponsor = False
    Else
        IncludeSponsor = GetValue(cINCSPONSOR)
    End If
End Property
Public Property Let IncludeSponsor(ByVal NewValue As Boolean)
    SetValue cINCSPONSOR, NewValue
End Property
Public Property Get IncludeBasicOnly() As Boolean
    If GetValue(cINCBASIC) = vbNullString Then
        IncludeBasicOnly = False
    Else
        IncludeBasicOnly = GetValue(cINCBASIC)
    End If
End Property
Public Property Let IncludeBasicOnly(ByVal NewValue As Boolean)
    SetValue cINCBASIC, NewValue
End Property
Public Property Get QuickSearch() As Boolean
    If GetValue(cQUKSEARCH) = vbNullString Then
        QuickSearch = False
    Else
        QuickSearch = GetValue(cQUKSEARCH)
    End If
End Property
Public Property Let QuickSearch(ByVal NewValue As Boolean)
    SetValue cQUKSEARCH, NewValue
End Property

Public Property Get LastRecordID() As String
    LastRecordID = GetValue(cLASTRECID)
End Property
Public Property Let LastRecordID(ByVal NewValue As String)
    SetValue cLASTRECID, NewValue
End Property
Public Property Get MaximumNumberOfRecords() As String
    MaximumNumberOfRecords = GetValue(cMAXNUMRECORDS)
End Property
Public Property Let MaximumNumberOfRecords(ByVal NewValue As String)
    SetValue cMAXNUMRECORDS, NewValue
End Property


Public Property Get SearchCHCS() As Boolean
    If GetValue(cCHCSSRH) = vbNullString Then
        SearchCHCS = False
    Else
        SearchCHCS = GetValue(cCHCSSRH)
    End If
End Property
Public Property Let SearchCHCS(ByVal NewValue As Boolean)
    SetValue cCHCSSRH, NewValue
End Property

Public Property Get GetAllPatients() As Boolean
    If GetValue(cGETALLPATS) = vbNullString Then
        GetAllPatients = False
    Else
        GetAllPatients = GetValue(cGETALLPATS)
    End If
End Property
Public Property Let GetAllPatients(ByVal NewValue As Boolean)
    SetValue cGETALLPATS, NewValue
End Property

Public Property Get ShowVIP() As Boolean
    If GetValue(cShowVIP) = vbNullString Then
        ShowVIP = False
    Else
        ShowVIP = GetValue(cShowVIP)
    End If
End Property
Public Property Let ShowVIP(ByVal NewValue As Boolean)
    SetValue cShowVIP, NewValue
End Property

Public Property Get SearchADS() As Boolean
    If GetValue(cShowADSFlag) = vbNullString Then
        SearchADS = False
    Else
        SearchADS = GetValue(cShowADSFlag)
    End If
End Property
Public Property Let SearchADS(ByVal NewValue As Boolean)
    SetValue cShowADSFlag, NewValue
End Property

'--- SCR 84911  Sherry Wang 1/6/2006
Public Property Get LoadUnverified() As Boolean
    If GetValue(cLoadUnverified) = vbNullString Then
        LoadUnverified = False
    Else
        LoadUnverified = GetValue(cLoadUnverified)
    End If
End Property
Public Property Let LoadUnverified(ByVal NewValue As Boolean)
    SetValue cLoadUnverified, NewValue
End Property


