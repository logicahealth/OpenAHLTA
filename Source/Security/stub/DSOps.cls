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
Attribute VB_Name = "DSOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_DSOps

Private Const ACTIVATE_KEY  As String = "ASKJHDSFIOUYCXOIVUZXCKBASDFBASOIXZYCVKASDBFMNEQ234KYF8S7D6F98AD7F87ASD9867FQWOEGHFASKDJHGVZXCYTV;"
Private Const CLASS_NAME    As String = "CHCSII_SecurityStubDS.DSOps"
Private m_bActivated        As Boolean
Private m_INIPath           As String
Private m_colSensitiveData  As New Collection
Private objDAS              As ICHCSII_DataAccess.ICHCSII_DAS
Private oConn               As CHCSII_CONN.Conn

'**************************
'SCR 18800
'Reading Data Elements into Array
Private Type ptSensitivityStub
   Type As String
   NCID As String
   Level As String
End Type
Dim SensArray() As ptSensitivityStub
'**************************

Private Sub Class_Initialize()
    m_bActivated = False

    Dim oShared As New CWShared
        m_INIPath = oShared.INIPath
        Set oConn = New Conn
        Set objDAS = oConn.CHCSII_DAS
    Set oShared = Nothing
    ReadSecurityStubInfo
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'SCR # 17079
'B. Fote
'29-Sep-2001
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'SCR 18800
'Retrieving Security Information from Database only once
'*******************************************

Private Sub ReadSecurityStubInfo()

'Reads from the Security_Stub_Data_Sensitivity CDR+ table
'  to pull in dummy security info

Dim sCurrentType As String
Dim strSQL As String
Dim mRS As ADODB.Recordset

Dim i As Integer
  
  strSQL = "SELECT Sensitivity_Type, Sensitivity_NCID, Sensitivity_Level " & _
            "FROM Security_Stub_Data_Sensitivity " & _
            "ORDER BY Sensitivity_Type, Sensitivity_NCID, Sensitivity_Level"
            
  
  Set mRS = objDAS.OpenRecordset(strSQL)
  
  If mRS.EOF And mRS.BOF Then
    MsgBox "There are no records in the Security_Stub_DataSensitivity table.", vbCritical, "Data Sensitivity"
    Set mRS = Nothing
  Else
    
    'Transfer into the Array
    ReDim SensArray(0)
    
    mRS.MoveFirst
    Do Until mRS.EOF
      ReDim Preserve SensArray(UBound(SensArray) + 1)
      
      SensArray(UBound(SensArray)).Type = UCase(mRS.Fields(0))
      SensArray(UBound(SensArray)).NCID = mRS.Fields(1)
      SensArray(UBound(SensArray)).Level = mRS.Fields(2)
      
      mRS.MoveNext
          
    Loop
    
    Set mRS = Nothing
    
    sCurrentType = ""
    For i = 1 To UBound(SensArray)
    
      If SensArray(i).Type <> sCurrentType Then
        
        sCurrentType = SensArray(i).Type
        AddSensitiveCollection sCurrentType
        
      End If
       
    Next i
    
  End If
  
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'SCR # 17079
'B. Fote
'29-Sep-2001
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'SCR 18800
'Patch 535.3
'Reading through the Array to Populate Sensitivity Collection
'*******************************************

Private Sub AddSensitiveCollection(sCurrentType As String)

  Dim i, j As Integer
  Dim sCurrentNCID As String
  Dim sCurrentLevel As String
  
  Dim objSensitive As Sensitive
  Dim objICHCSII_Sensitive As ICHCSII_Sensitive
  Dim colDest As New Collection
  
  sCurrentNCID = ""

  For i = 1 To UBound(SensArray)
    
    If SensArray(i).Type = sCurrentType Then
    
      If SensArray(i).NCID <> sCurrentNCID Then
        
        sCurrentNCID = SensArray(i).NCID
        Set objSensitive = New Sensitive
        objSensitive.SetNCID sCurrentNCID
        Set objICHCSII_Sensitive = objSensitive
        
        j = i
                
        sCurrentLevel = ""
          
        Do Until SensArray(j).NCID <> sCurrentNCID
        
          If SensArray(j).Type <> sCurrentType Then Exit Do
          
          objICHCSII_Sensitive.AddSensitivity SensArray(j).Level
          j = j + 1
          
          If j > UBound(SensArray) Then Exit Do

        Loop
            
        colDest.Add objICHCSII_Sensitive, objICHCSII_Sensitive.GetNCID()
        
      End If
    
    End If
    
  Next i
  
  m_colSensitiveData.Add colDest, sCurrentType
  
End Sub


Private Sub Class_Terminate()
    Set objDAS = Nothing
End Sub

Private Function ICHCSII_DSOps_Activate(ActivateKey As String) As Boolean
    If (ActivateKey = ACTIVATE_KEY) Then
        m_bActivated = True
        ICHCSII_DSOps_Activate = True
    Else
        m_bActivated = False
        ICHCSII_DSOps_Activate = False
    End If
End Function

Private Function ICHCSII_DSOps_GetSensitivitiesForNCIDs(NCIDs() As String, DomainNCID As String) As CHCSII_ClientSecurity.ICHCSII_Sensitive()

If Not (m_bActivated) Then
    Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
End If

Dim iNumItems As Integer
Dim iNumSensitivities As Integer
Dim i As Integer
Dim j As Integer
Dim colSensitiveDomain As Collection
Dim objICHCSII_Sensitive As ICHCSII_Sensitive
Dim objSensitive As Sensitive
Dim sSensitivities() As String
Dim colSensitiveItems As New Collection
Dim sensitiveItems() As ICHCSII_Sensitive

On Error Resume Next

iNumItems = UBound(NCIDs)
If Err Then
    Err.Clear
    Exit Function
Else
    Set colSensitiveDomain = m_colSensitiveData.Item(DomainNCID)
    If Err Then
        On Error GoTo 0
        Err.Raise SEC_DOMAIN_NOT_FOUND, CLASS_NAME, "Domain " + DomainNCID + " not found."
    End If
    
    For i = 1 To iNumItems
        Set objSensitive = New Sensitive
        objSensitive.SetNCID NCIDs(i)
        Set objICHCSII_Sensitive = colSensitiveDomain.Item(NCIDs(i))
        If Err Then
            'Item Not Sensitive
            Err.Clear
        Else
            sSensitivities = objICHCSII_Sensitive.GetSensitivities()
            iNumSensitivities = UBound(sSensitivities)
            If Err Then
                'No sensitivity levels found
                Err.Clear
            Else
                Set objICHCSII_Sensitive = objSensitive
                'Add sensitivity levels to sensitive data item
                For j = 1 To iNumSensitivities
                    objICHCSII_Sensitive.AddSensitivity sSensitivities(j)
                Next
            End If
        End If
        colSensitiveItems.Add objSensitive
    Next
End If

On Error GoTo 0

If colSensitiveItems.Count > 0 Then
    ReDim sensitiveItems(1 To colSensitiveItems.Count)
    For i = 1 To colSensitiveItems.Count
        Set sensitiveItems(i) = colSensitiveItems.Item(i)
    Next
End If

Set colSensitiveItems = Nothing
ICHCSII_DSOps_GetSensitivitiesForNCIDs = sensitiveItems

End Function

Private Sub ICHCSII_DSOps_SetSensitivities(SensitiveData() As CHCSII_ClientSecurity.ICHCSII_Sensitive, DomainNCID As String)

If Not (m_bActivated) Then
    Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
End If

Dim iNumItems As Integer
Dim iNumSensitivities As Integer
Dim i As Integer
Dim j As Integer
Dim colSensitiveDomain As Collection
Dim objICHCSII_Sensitive As ICHCSII_Sensitive
Dim sSensitivities() As String

On Error Resume Next

iNumItems = UBound(SensitiveData)
If Err Then
    Err.Clear
    Exit Sub
Else
    Set colSensitiveDomain = m_colSensitiveData.Item(DomainNCID)
    If Err Then
        On Error GoTo 0
        Err.Raise SEC_DOMAIN_NOT_FOUND, CLASS_NAME, "Domain " + DomainNCID + " not found."
    End If
    
    For i = 1 To iNumItems
        Set objICHCSII_Sensitive = colSensitiveDomain.Item(SensitiveData(i).GetNCID())
        If Err Then
            'Item Not Sensitive
            Err.Clear
        Else
            sSensitivities = objICHCSII_Sensitive.GetSensitivities()
            iNumSensitivities = UBound(sSensitivities)
            If Err Then
                'No sensitivity levels found
                Err.Clear
            Else
                'Add sensitivity levels to sensitive data item
                For j = 1 To iNumSensitivities
                    SensitiveData(i).AddSensitivity sSensitivities(j)
                Next
            End If
        End If
    Next
End If

End Sub


