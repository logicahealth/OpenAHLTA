VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperCHCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperNonProviderList

Private mobjMapperGeneric As iMapperNonProviderList

Private Sub Class_Initialize()
    Dim objOsh As iLegacyOsh
    Dim objMapper As MapperGeneric
    
    On Error GoTo ErrHandler
    
    Set mobjMapperGeneric = New MapperGeneric
    Set objOsh = New LegacyOshCDR
    Set objMapper.LegacyOsh = objOsh
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.Initialize", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub iMapperNonProviderList_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    Call mobjMapperGeneric.Init(Support, Cache)
End Sub

Private Sub iMapperNonProviderList_SaveList(NPParams As DL_Support.iParameters)
    Call mobjMapperGeneric.SaveList(NPParams)
End Sub

Private Function iMapperNonProviderList_Selected(NPParams As DL_Support.iParameters) As DTOs.Providers
    Set iMapperNonProviderList_Selected = mobjMapperGeneric.Selected(NPParams)
End Function

