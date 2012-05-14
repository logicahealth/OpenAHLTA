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
Attribute VB_Name = "EncProviders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE_NAME = "EncProviders"

Private mcProviders As Collection
Private moShared As CWShared

Private Sub Class_Initialize()
  Set mcProviders = New Collection
  Set moShared = New CWShared
End Sub

Private Sub Class_Terminate()
   Set mcProviders = Nothing
   Set moShared = Nothing
End Sub

Public Function Exists(ByVal sNCID As String, Optional ByRef oProvider As EncProvider) As Boolean
On Error GoTo ErrorHandler

   Set oProvider = Find(sNCID)
   Exists = CBool(Not oProvider Is Nothing)
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "Find"
   
End Function
Public Function Find(ByVal sProviderNCID As String, Optional ByRef lPosition As Long) As EncProvider

On Error GoTo ErrorHandler

Dim lIdx As Long

   For lIdx = 1 To mcProviders.Count
      If mcProviders.Item(lIdx).NCID = sProviderNCID Then
         Set Find = mcProviders.Item(lIdx)
         lPosition = lIdx
         Exit For
      End If
   Next lIdx
      
ErrorHandler:
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "Find"
                
End Function
Public Function Add(ByRef oProvider As EncProvider, ByVal sProviderNCID As String, Optional ByVal lBefore = 0) As EncProvider
   If lBefore > 0 And mcProviders.Count > 0 Then
      mcProviders.Add oProvider, sProviderNCID, lBefore
   Else
      mcProviders.Add oProvider, sProviderNCID
   End If
End Function

Public Function Count() As Long
  Count = mcProviders.Count
End Function

Public Function Item(ByVal Key As Variant) As EncProvider
Attribute Item.VB_UserMemId = 0
  Set Item = mcProviders.Item(Key)
End Function

Public Sub Remove(ByVal Key As Variant)
   mcProviders.Remove Key
End Sub

'Remove a provider from the providers collection
Friend Sub RemoveProvider(ByVal sNCID As String)

On Error GoTo ErrorHandler

Dim oProvider As EncProvider
   
   If Not Exists(sNCID, oProvider) Then Exit Sub
   
   If moShared.IsAppMode(modeCDR) And Not moShared.IsAppMode(modeITT) Then
      'Don't need to worry about state in CDR mode so just remove it
      Remove sNCID
   Else
      If oProvider.State = rsNew Or oProvider.State = rsRemove Then
         Remove sNCID
      Else
         oProvider.State = rsDelete
      End If
   End If
   
ErrorHandler:
   Set oProvider = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "RemoveProvider"
                             
End Sub

'Remove deleted providers from the providers collection
Friend Sub RemoveProviders()

On Error GoTo ErrorHandler

Dim lIdx As Long
Dim oProvider As EncProvider

   For lIdx = mcProviders.Count To 1 Step -1
      Set oProvider = mcProviders(lIdx)
      If oProvider.State = rsRemove Then
         mcProviders.Remove lIdx
      End If
   Next
   
ErrorHandler:
   Set oProvider = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "RemoveProviders"
                
End Sub
'Added for backwards compatibility so that modules that use the providers collection will still work
'This can be removed once all other modules are updated to use the EncProviders object
Public Property Get Providers() As Collection
   Set Providers = mcProviders
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcProviders.[_NewEnum]
End Property

Public Function IsDirty() As Boolean

On Error GoTo ErrorHandler

Dim oProv As EncProvider

   IsDirty = False
   
   If (moShared.IsAppMode(modeCHCSI_GUI) Or moShared.IsAppMode(modeCDR)) _
      And Len(moShared.CmdLineSwitches("HIPAA837")) > 0 Then
         
      For Each oProv In mcProviders
         If oProv.IsDirty() Then
            IsDirty = True
            Exit For
         End If
      Next
      
   End If
   
ErrorHandler:
   Set oProv = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "IsDirty"
  
End Function








 
 
