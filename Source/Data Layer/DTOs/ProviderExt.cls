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
Attribute VB_Name = "ProviderExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Application reference item in tree structure.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "ProviderExt"
Private Const OBJECT_VERSION As String = "1.0"

Private msNCID                   As String
Private msFirstName              As String
Private msLastName               As String
Private msMiddleName             As String
Private msFullName               As String
Private msPhone1                 As String
Private msPhone2                 As String
Private msProviderIEN            As String
Private msSignatureClass         As String
Private msUserIEN                As String
Private msSSN                    As String
Private msSpecialityCode         As String
Private msDefaultClinicNCID      As String
Private msFacilityNCID           As String

Private Function iObjectPersist_Copy() As Object
Dim Item As ProviderExt
  On Error GoTo errHandler
  
  Set Item = New ProviderExt
  
  With Item
    .NCID = msNCID
    .FirstName = msFirstName
    .LastName = msLastName
    .MiddleName = msMiddleName
    .FullName = msFullName
    .Phone1 = msPhone1
    .Phone2 = msPhone2
    .ProviderIEN = msProviderIEN
    .SignatureClass = msSignatureClass
    .UserIEN = msUserIEN
    .SSN = msSSN
    .SpecialityCode = msSpecialityCode
    .DefaultClinicNCID = msDefaultClinicNCID
    .FacilityNCID = msFacilityNCID
  End With
  
  Set iObjectPersist_Copy = Item
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "DTOsProviderExt.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Sub iObjectPersist_Persist(ByRef Writer As ObjectPersist.iWriter)
  On Error GoTo errHandler
  
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    .WriteProperty "NCID", msNCID
    .WriteProperty "FirstName", msFirstName
    .WriteProperty "LastName", msLastName
    .WriteProperty "MiddleName", msMiddleName
    .WriteProperty "FullName", msFullName
    .WriteProperty "Phone1", msPhone1
    .WriteProperty "Phone2", msPhone2
    .WriteProperty "ProviderIEN", msProviderIEN
    .WriteProperty "SignatureClass", msSignatureClass
    .WriteProperty "UserIEN", msUserIEN
    .WriteProperty "SSN", msSSN
    .WriteProperty "SpecialityCode", msSpecialityCode
    .WriteProperty "DefaultClinicNCID", msDefaultClinicNCID
    .WriteProperty "FacilityNCID", msFacilityNCID
  End With
  Exit Sub
    
errHandler:
    Err.Raise Err.Number, "DTOsProviderExt.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function iObjectPersist_Restore(ByRef Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo errHandler
  
  With Reader
    msNCID = .ReadProperty("NCID")
    msFirstName = .ReadProperty("FirstName")
    msLastName = .ReadProperty("LastName")
    msMiddleName = .ReadProperty("MiddleName")
    msFullName = .ReadProperty("FullName")
    msPhone1 = .ReadProperty("Phone1")
    msPhone2 = .ReadProperty("Phone2")
    msProviderIEN = .ReadProperty("ProviderIEN")
    msSignatureClass = .ReadProperty("SignatureClass")
    msUserIEN = .ReadProperty("UserIEN")
    msSSN = .ReadProperty("SSN")
    msSpecialityCode = .ReadProperty("SpecialityCode")
    msDefaultClinicNCID = .ReadProperty("DefaultClinicNCID")
    msFacilityNCID = .ReadProperty("FacilityNCID")
  End With

  iObjectPersist_Restore = Not Failed
  
  Exit Function
  
errHandler:
  Failed = True
  Resume Next

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Property Get NCID() As String
    NCID = msNCID
End Property
Public Property Let NCID(ByVal NewValue As String)
    msNCID = NewValue
End Property

Public Property Get FirstName() As String
    FirstName = msFirstName
End Property
Public Property Let FirstName(ByVal NewValue As String)
    msFirstName = NewValue
End Property

Public Property Get LastName() As String
    LastName = msLastName
End Property
Public Property Let LastName(ByVal NewValue As String)
    msLastName = NewValue
End Property

Public Property Get MiddleName() As String
    MiddleName = msMiddleName
End Property
Public Property Let MiddleName(ByVal NewValue As String)
    msMiddleName = NewValue
End Property

Public Property Get FullName() As String
    FullName = msFullName
End Property
Public Property Let FullName(ByVal NewValue As String)
    msFullName = NewValue
End Property

Public Property Get Phone1() As String
    Phone1 = msPhone1
End Property
Public Property Let Phone1(ByVal NewValue As String)
    msPhone1 = NewValue
End Property

Public Property Get Phone2() As String
    Phone2 = msPhone2
End Property
Public Property Let Phone2(ByVal NewValue As String)
    msPhone2 = NewValue
End Property

Public Property Get ProviderIEN() As String
    ProviderIEN = msProviderIEN
End Property
Public Property Let ProviderIEN(ByVal NewValue As String)
    msProviderIEN = NewValue
End Property

Public Property Get SignatureClass() As String
    SignatureClass = msSignatureClass
End Property
Public Property Let SignatureClass(ByVal NewValue As String)
    msSignatureClass = NewValue
End Property

Public Property Get UserIEN() As String
    UserIEN = msUserIEN
End Property
Public Property Let UserIEN(ByVal NewValue As String)
    msUserIEN = NewValue
End Property

Public Property Get SSN() As String
    SSN = msSSN
End Property
Public Property Let SSN(ByVal NewValue As String)
    msSSN = NewValue
End Property

Public Property Get SpecialityCode() As String
    SpecialityCode = msSpecialityCode
End Property
Public Property Let SpecialityCode(ByVal NewValue As String)
    msSpecialityCode = NewValue
End Property

Public Property Get DefaultClinicNCID() As String
    DefaultClinicNCID = msDefaultClinicNCID
End Property
Public Property Let DefaultClinicNCID(ByVal NewValue As String)
    msDefaultClinicNCID = NewValue
End Property

Public Property Get FacilityNCID() As String
    FacilityNCID = msFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal NewValue As String)
    msFacilityNCID = NewValue
End Property


