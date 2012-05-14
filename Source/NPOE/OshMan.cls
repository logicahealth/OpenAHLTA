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
Attribute VB_Name = "OshMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private lResult As Long
Private sName As String
Private iCount As Integer
Private mobjConfig As Config

Friend Function GetOshItems(sSearchClass As String, sName As String, sOwner As String) As Collection

    Dim colItems As Collection
    Dim objShared As New CWShared
    Dim objGemsDas As ICHCSII_DAS
    Dim objRecSet As New Recordset
    Dim sSQL As String
    Dim sLocalValue As String
    Dim sLocalArray() As String
    Dim iCount As Integer
    On Error GoTo ErrHandler
    
    ' Retrieve from OSH tables directly
    Set objGemsDas = SetCHCSII_DAS
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
        Set GetOshItems = colItems
    Else
        Set colItems = Nothing
        Set GetOshItems = Nothing
    End If

Exit Function

ErrHandler:
    MsgBox "Error getting osh Bins", vbCritical, "Non-Provider Order Entry - OSH Error"
    Set GetOshItems = Nothing
    Exit Function
End Function
Friend Property Let ConfigObject(objConfig As Object)
'ADDProviders change
'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:09 PM
'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode
Set mobjConfig = objConfig
End Property


Friend Function CreateOshItems(sClass As String, sType As String, sName As String, sOwner As String, colItems As Collection) As Boolean
'ADDProviders change
    'R. Foster SCR #32633;   Developer: Robert Foster 02/19/2003 04:09 PM
    'R. Foster Note: Function Changed:   ProviderList.frm: Display; cmdClear_Click; cmdSearch_Click; cboProv_Click  clsNonProvider.cls: SearchAdditionalProviders; getAdditionalProvidersList  OshMan.cls: CreateOshItems;    Impact of Change: Enable MRU List in Additional Providers in CHCS_GUI mode
    Dim objShared As New CWShared
    Dim objGemsDas As New ICHCSII_DAS
    Dim objRecSet As New Recordset
    Dim sSQL As String
    Dim strBuildValue As String
    On Error GoTo ErrHandler
    
    ' Retrieve from OSH tables directly
    Set objGemsDas = SetCHCSII_DAS
    For iCount = colItems.Count To 1 Step -1
        sSQL = "insert into osh_ini_settings (apownerncid, apsection, apkey, apvalue) values " & _
        "(" & sOwner & ",'" & sClass & "','" & sType & "','" & colItems(iCount) & "')"
        objGemsDas.ExecuteSQL sSQL
    Next

    CreateOshItems = True
Exit Function

ErrHandler:
'    MsgBox "Error Saving Bin Items"
'    MsgBox "Error Saving Bin Items", vbCritical, "Non-Provider Order Entry - Error Saving To OSH"
    MsgBox "Error Saving Bin Items, please see your System Administrator", vbCritical, "Non-Provider Order Entry - Error Saving To OSH"
    CreateOshItems = False
    Exit Function
End Function

Friend Function DeleteOshItems(sSearchClass As String, sName As String, sOwner As String) As Boolean

    Dim objShared As New CWShared
    Dim objGemsDas As New ICHCSII_DAS
    Dim objRecSet As New Recordset
    Dim sSQL As String
    
    On Error GoTo ErrHandler
        
    ' Retrieve from OSH tables directly
    Set objGemsDas = SetCHCSII_DAS
    sSQL = "Delete from osh_ini_settings where " & _
        "apsection = '" & sSearchClass & "' and apkey = '" & sName & _
        "' and apownerncid = " & sOwner
    objGemsDas.ExecuteSQL (sSQL)
    
    DeleteOshItems = True
Exit Function

ErrHandler:
'    MsgBox "Error getting osh Bins"
    MsgBox "Error getting osh Bins - Please see your system administrator.", vbCritical, "Non-Provider Order Entry - Error Reading Osh"
    DeleteOshItems = False
    Exit Function
End Function

