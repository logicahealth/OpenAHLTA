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
Attribute VB_Name = "ModuleConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Defines an application server (module) in CHCSII.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "OLEServer"
Private Const OBJECT_VERSION As String = "1.0"

Private mAccessPriv As String                 'Name of priv. required to use application server
Private mApplicationNCID As String            'Unique application idenifier (Numeric Concept ID)
Private mApplicationType As String            'PATIENT_SEARCH_APP, PATIENT_ID_APP, CLINICAL_APP, SUMMARY_APP, TOOLS_APP, SUPPORT_APP
Private mButtonCaption As String              'Text used in any button that represents the application server
Private mCommandLineSwitch As String          'Switch require to use this application server
Private mHasPatientSummary As Boolean         'Application server has a patient summary view
Private mHasPrint As Boolean                  'Application server has print options
Private mHasPrintPreview As Boolean           'Application server has print preview options
Private mHasPropertyView As Boolean           'Application server has a property view (settings)
Private mImageName As String                  'Image name used to represent this application server
Private mMenuCaption As String                'Test used in any menu that represents the application server
Private mPatientRequired As Integer           'True if a Patient must be selected to use this application server
Private mRegisterName As String               'Component name where application server is located (DLL or OCX)
Private mSelected As Boolean                  'True if module is available and false if it can not be used
Private mServerName As String                 'Class name that can be used to create an instance of the application server
Private mToolTip As String                    'Text used for any tooltip related to this applicatioon server

Private Function iObjectPersist_Copy() As Object
Dim Item As ModuleConfig

  Set Item = New ModuleConfig

  With Item
    .AccessPriv = mAccessPriv
    .ApplicationNCID = mApplicationNCID
    .ApplicationType = mApplicationType
    .ButtonCaption = mButtonCaption
    .CommandLineSwitch = mCommandLineSwitch
    .HasPatientSummary = mHasPatientSummary
    .HasPrint = mHasPrint
    .HasPrintPreview = mHasPrintPreview
    .HasPropertyView = mHasPropertyView
    .ImageName = mImageName
    .MenuCaption = mMenuCaption
    .PatientRequired = mPatientRequired
    .RegisterName = mRegisterName
    .Selected = mSelected
    .ServerName = mServerName
    .ToolTip = mToolTip
  End With

  Set iObjectPersist_Copy = Item

End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION

    .WriteProperty "AccessPriv", mAccessPriv
    .WriteProperty "ApplicationNCID", mApplicationNCID
    .WriteProperty "ApplicationType", mApplicationType
    .WriteProperty "ButtonCaption", mButtonCaption
    .WriteProperty "CommandLineSwitch", mCommandLineSwitch
    .WriteProperty "HasPatientSummary", mHasPatientSummary
    .WriteProperty "HasPrint", mHasPrint
    .WriteProperty "HasPrintPreview", mHasPrintPreview
    .WriteProperty "HasPropertyView", mHasPropertyView
    .WriteProperty "ImageName", mImageName
    .WriteProperty "MenuCaption", mMenuCaption
    .WriteProperty "PatientRequired", mPatientRequired
    .WriteProperty "RegisterName", mRegisterName
    .WriteProperty "Selected", mSelected
    .WriteProperty "ServerName", mServerName
    .WriteProperty "ToolTip", mToolTip
  End With
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler

  With Reader
    mAccessPriv = .ReadProperty("AccessPriv")
    mApplicationNCID = .ReadProperty("ApplicationNCID")
    mApplicationType = .ReadProperty("ApplicationType")
    mButtonCaption = .ReadProperty("ButtonCaption")
    mCommandLineSwitch = .ReadProperty("CommandLineSwitch")
    mHasPatientSummary = .ReadProperty("HasPatientSummary")
    mHasPrint = .ReadProperty("HasPrint")
    mHasPrintPreview = .ReadProperty("HasPrintPreview")
    mHasPropertyView = .ReadProperty("HasPropertyView")
    mImageName = .ReadProperty("ImageName")
    mMenuCaption = .ReadProperty("MenuCaption")
    mPatientRequired = .ReadProperty("PatientRequired")
    mRegisterName = .ReadProperty("RegisterName")
    mSelected = .ReadProperty("Selected")
    mServerName = .ReadProperty("ServerName")
    mToolTip = .ReadProperty("ToolTip")
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

Public Property Get AccessPriv() As String
  AccessPriv = mAccessPriv
End Property

Public Property Let AccessPriv(ByVal NewValue As String)
  mAccessPriv = NewValue
End Property

Public Property Get ApplicationNCID() As String
  ApplicationNCID = mApplicationNCID
End Property

Public Property Let ApplicationNCID(ByVal NewValue As String)
  mApplicationNCID = NewValue
End Property

Public Property Get ApplicationType() As String
  ApplicationType = mApplicationType
End Property

Public Property Let ApplicationType(ByVal NewValue As String)
  mApplicationType = NewValue
End Property

Public Property Get ButtonCaption() As String
  ButtonCaption = mButtonCaption
End Property

Public Property Let ButtonCaption(ByVal NewValue As String)
  mButtonCaption = NewValue
End Property

Public Property Get CommandLineSwitch() As String
  CommandLineSwitch = mCommandLineSwitch
End Property

Public Property Let CommandLineSwitch(ByVal NewValue As String)
  mCommandLineSwitch = NewValue
End Property

Public Property Get HasPatientSummary() As Boolean
  HasPatientSummary = mHasPatientSummary
End Property

Public Property Let HasPatientSummary(ByVal NewValue As Boolean)
  mHasPatientSummary = NewValue
End Property

Public Property Get HasPrint() As Boolean
  HasPrint = mHasPrint
End Property

Public Property Let HasPrint(ByVal NewValue As Boolean)
  mHasPrint = NewValue
End Property

Public Property Get HasPrintPreview() As Boolean
  HasPrintPreview = mHasPrintPreview
End Property

Public Property Let HasPrintPreview(ByVal NewValue As Boolean)
  mHasPrintPreview = NewValue
End Property

Public Property Get HasPropertyView() As Boolean
  HasPropertyView = mHasPropertyView
End Property

Public Property Let HasPropertyView(ByVal NewValue As Boolean)
  mHasPropertyView = NewValue
End Property

Public Property Get ImageName() As String
  ImageName = mImageName
End Property

Public Property Let ImageName(ByVal NewValue As String)
  mImageName = NewValue
End Property

Public Property Get MenuCaption() As String
  MenuCaption = mMenuCaption
End Property

Public Property Let MenuCaption(ByVal NewValue As String)
  mMenuCaption = NewValue
End Property

Public Property Get PatientRequired() As Integer
  PatientRequired = mPatientRequired
End Property

Public Property Let PatientRequired(ByVal NewValue As Integer)
  mPatientRequired = NewValue
End Property

Public Property Get RegisterName() As String
  RegisterName = mRegisterName
End Property

Public Property Let RegisterName(ByVal NewValue As String)
  mRegisterName = NewValue
End Property

Public Property Get Selected() As Boolean
  Selected = mSelected
End Property

Public Property Let Selected(ByVal NewValue As Boolean)
  mSelected = NewValue
End Property

Public Property Get ServerName() As String
  ServerName = mServerName
End Property

Public Property Let ServerName(ByVal NewValue As String)
  mServerName = NewValue
End Property

Public Property Get ToolTip() As String
  ToolTip = mToolTip
End Property

Public Property Let ToolTip(ByVal NewValue As String)
  mToolTip = NewValue
End Property

