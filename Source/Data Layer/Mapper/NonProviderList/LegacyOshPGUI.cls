VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LegacyOshPGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NOTE:  This class provides support for retrieving NonProvider lists that were stored in the original
'NPOE class.  The current DataLayer Mapper functionality does not store values via these methods.
'At some time the support here should be dropped and rely on DataLayer OSH to store the settings.

Option Explicit

Implements iLegacyOsh

Private Property Set iLegacyOsh_Conn(RHS As CHCSII_CONN.Conn)
'Don't need it
End Property

Private Function iLegacyOsh_GetOshItems(ByVal sSearchClass As String, ByVal sName As String, ByVal sOwner As String) As Collection
'No Legacy support needed, just return an emtpy collection

    Dim colItems As Collection
    
    On Error GoTo ErrHandler

    Set colItems = New Collection
    
    Set iLegacyOsh_GetOshItems = colItems

    Set colItems = Nothing
Exit Function

ErrHandler:
    Err.Raise Err.Number, "LegacyOshPGUI.GetOshItems", Err.Description, Err.HelpFile, Err.HelpContext
    Set iLegacyOsh_GetOshItems = Nothing
    Exit Function
End Function

