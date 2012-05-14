VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Template_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private msTemplateID As String
Private msTemplateName As String
Private msPrefix As String
Private msFolderID As String
Private msFolderName As String
Private msOwner As String
Private msType As TemplateOwnerTypeEnum



Public Property Get TemplateID() As String
    TemplateID = msTemplateID
End Property

Public Property Let TemplateID(ByVal RHV As String)
    msTemplateID = RHV
End Property

Public Property Get TemplateName() As String
    TemplateName = msTemplateName
End Property

Public Property Let TemplateName(ByVal RHV As String)
    msTemplateName = RHV
End Property

Public Property Get FolderID() As String
    FolderID = msFolderID
End Property
Public Property Let FolderID(ByVal RHV As String)
    msFolderID = RHV
End Property


Public Property Get FolderName() As String
    FolderName = msFolderName
End Property

Public Property Let FolderName(ByVal RHV As String)
    msFolderName = RHV
End Property

Public Property Get TemplatePrefix() As String
    TemplatePrefix = msPrefix
End Property

Public Property Let TemplatePrefix(ByVal RHV As String)
    msPrefix = RHV
End Property

Public Property Get Owner() As String
    Owner = msOwner
End Property

Public Property Let Owner(ByVal RHV As String)
    msOwner = RHV
End Property

Public Property Get OwnerType() As TemplateOwnerTypeEnum
    OwnerType = msType
End Property

Public Property Let OwnerType(ByVal RHV As TemplateOwnerTypeEnum)
    msType = RHV
End Property
