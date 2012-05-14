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
Attribute VB_Name = "LegacyOshTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NOTE:  This class provides support for retrieving NonProvider lists that were stored in the original
'NPOE class.  The current DataLayer Mapper functionality does not store values via these methods.
'At some time the support here should be dropped and rely on DataLayer OSH to store the settings.

Option Explicit

Implements iLegacyOsh

Private mConn As Conn

Private Property Set iLegacyOsh_Conn(RHS As CHCSII_CONN.Conn)
    Set mConn = RHS
End Property

Private Function iLegacyOsh_GetOshItems(ByVal sSearchClass As String, ByVal sName As String, ByVal sOwner As String) As Collection
Dim colItems As Collection
Dim objGemsDas As ICHCSII_DAS
Dim objRecSet As Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
'

    Set objRecSet = New Recordset
    
    ' Retrieve from OSH tables directly
    Set objGemsDas = mConn.CHCSII_DAS
    sSQL = "Select apkey, apvalue from osh_ini_settings where " & _
        "apsection = '" & sSearchClass & "' and apkey = '" & sName & _
        "' and apownerncid = " & sOwner
    Set objRecSet = objGemsDas.OpenRecordset(sSQL)
    If Not objRecSet.EOF And Not objRecSet.BOF Then
        Set colItems = New Collection
        objRecSet.MoveFirst
        Do While Not objRecSet.EOF
            colItems.Add objRecSet(1).Value & "", objRecSet(1).Value & ""
            objRecSet.MoveNext
        Loop
        Set iLegacyOsh_GetOshItems = colItems
    Else
        Set colItems = Nothing
        Set iLegacyOsh_GetOshItems = Nothing
    End If
    
    objRecSet.Close
    Set objRecSet = Nothing
    Set colItems = Nothing
    
Exit Function

ErrHandler:
    Err.Raise Err.Number, "LegacyOshTheater.GetOshItems", Err.Description, Err.HelpFile, Err.HelpContext
    Set iLegacyOsh_GetOshItems = Nothing
    Exit Function

End Function

