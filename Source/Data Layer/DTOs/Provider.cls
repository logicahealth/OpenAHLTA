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
Attribute VB_Name = "Provider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Application reference item in tree structure.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "Provider"
Private Const OBJECT_VERSION As String = "1.0"

Private mstrIEN As String               'Provider IEN
Private mstrNCID As String              'Provider NCID
Private mstrName As String              'Provider Name
Private mintSigClass As Integer         'Provider Signature Class
Private mstrLogonName As String         'Provider Logon Name
Private mstrRoles As String             'Provider Roles
Private mcolWorkloads As Workloads      'Collection Provider Workloads
Private mcolExtendedData As ProviderExt 'Collection of Provider Extended Data (Name, Address, Speciality ...)

Private Function iObjectPersist_Copy() As Object
Dim Item As Provider
Dim objPersist As iObjectPersist

  On Error GoTo ErrHandler
  
  Set Item = New Provider
  
  With Item
    .IEN = mstrIEN
    .NCID = mstrNCID
    .Name = mstrName
    .SignatureClass = mintSigClass
    .LogonName = mstrLogonName
    .Roles = mstrRoles
    
    Set objPersist = New Workloads
    Set .ProviderWorkloads = objPersist.Copy
    
    Set objPersist = New ProviderExt
    Set .ProviderExtendedData = objPersist.Copy
  End With
  
  Set iObjectPersist_Copy = Item

  Set Item = Nothing

  Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "Provider.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  On Error GoTo ErrHandler
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    .WriteProperty "IEN", mstrIEN
    .WriteProperty "NCID", mstrNCID
    .WriteProperty "Name", mstrName
    .WriteProperty "SignatureClass", mintSigClass
    .WriteProperty "LogonName", mstrLogonName
    .WriteProperty "Roles", mstrRoles
    
    .WriteObject "Workloads", mcolWorkloads
    .WriteObject "ExtendedData", mcolExtendedData
  End With
  Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsProvider.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mstrIEN = .ReadProperty("IEN")
    mstrNCID = .ReadProperty("NCID")
    mstrName = .ReadProperty("Name")
    mintSigClass = .ReadProperty("SignatureClass")
    mstrLogonName = .ReadProperty("LogonName")
    mstrRoles = .ReadProperty("Roles")
    
    .ReadObject "Workloads", ProviderWorkloads
    .ReadObject "ExtendedData", ProviderExtendedData
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


Public Property Get IEN() As String
    IEN = mstrIEN
End Property
Public Property Let IEN(ByVal NewValue As String)
    mstrIEN = NewValue
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
Public Property Get SignatureClass() As Integer
    SignatureClass = mintSigClass
End Property
Public Property Let SignatureClass(ByVal NewValue As Integer)
    mintSigClass = NewValue
End Property
Public Function IsSignable() As Boolean
    Select Case mintSigClass
        Case 2, 3, 4
            IsSignable = True
        Case Else
            IsSignable = False
    End Select
End Function

Public Property Get LogonName() As String
    LogonName = mstrLogonName
End Property
Public Property Let LogonName(ByVal NewValue As String)
    mstrLogonName = NewValue
End Property

Public Property Get Roles() As String
    Roles = mstrRoles
End Property
Public Property Let Roles(ByVal NewValue As String)
    mstrRoles = NewValue
End Property

Public Property Get ProviderWorkloads() As Workloads
  If mcolWorkloads Is Nothing Then
    Set mcolWorkloads = New Workloads
  End If
  
  Set ProviderWorkloads = mcolWorkloads
End Property

Public Property Set ProviderWorkloads(ByVal NewValue As Workloads)
  Set mcolWorkloads = NewValue
End Property

Public Property Get ProviderExtendedData() As ProviderExt
  If mcolExtendedData Is Nothing Then
    Set mcolExtendedData = New ProviderExt
  End If
  
  Set ProviderExtendedData = mcolExtendedData
End Property

Public Property Set ProviderExtendedData(ByVal NewValue As ProviderExt)
  Set mcolExtendedData = NewValue
End Property

