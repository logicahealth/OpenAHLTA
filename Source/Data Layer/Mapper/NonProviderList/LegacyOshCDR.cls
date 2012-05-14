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
Attribute VB_Name = "LegacyOshCDR"
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
'Dont Need it
End Property

Private Function iLegacyOsh_GetOshItems(ByVal sSearchClass As String, ByVal sName As String, ByVal sOwner As String) As Collection

    Dim objOsh As Storehouse
    Dim objItem As OSHATLLib.BinItem
    Dim objBin As OSHATLLib.Bin
    Dim colItems As Collection
    
    On Error GoTo ErrHandler
        
        Set objOsh = New Storehouse
        
        ' Get those OSH records using the storehouse object
        objOsh.SearchClass = sSearchClass
        objOsh.SearchName = sName
        objOsh.SearchOwner = sOwner
        objOsh.Load True
        
        If objOsh.Bins.Count > 0 Then
            Set colItems = New Collection
            For Each objBin In objOsh.Bins
                For Each objItem In objBin.BinItems
                    colItems.Add objItem.Text, objItem.Name
                Next
            Next
            Set iLegacyOsh_GetOshItems = colItems
        Else
            Set colItems = Nothing
            Set iLegacyOsh_GetOshItems = Nothing
        End If

Exit Function

ErrHandler:
    Err.Raise Err.Number, "LegacyOshCDR.GetOshItems", Err.Description, Err.HelpFile, Err.HelpContext
    Set iLegacyOsh_GetOshItems = Nothing
    Exit Function

End Function
