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
Attribute VB_Name = "CPKI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements CHCSII_ClientSecurity.ICHCSII_PKI


Private Sub Class_Initialize()
    If goSecSQL Is Nothing Then
        Set goSecSQL = New GEMS_SQL.SqlOps
        Dim oSecSQL2 As ISec_SQL.ISecSqlOps
        Set oSecSQL2 = goSecSQL
        oSecSQL2.Connect   ' swicthes to "security user" schema
        Set oSecSQL2 = Nothing
    End If
End Sub


Private Function ICHCSII_PKI_GenerateKey(ByVal user_Ncid As String, ByVal sPassword As String, oUser As CHCSII_ClientSecurity.ICHCSII_User) As Boolean

End Function



Private Function ICHCSII_PKI_Sign(ByVal TextToSign As String, _
                                  ByVal sPassword As String, _
                                        oUser As CHCSII_ClientSecurity.ICHCSII_User) As String


'__ sign detached
On Error GoTo ErrHandler
    Dim sComment As String
    Dim bufferOut As String
    Dim bufferoutLen As Long
    bufferoutLen = 3000
    bufferOut = Space(bufferoutLen)
    
    Dim sUserID As String
    sUserID = GetUserId(oUser)
    
    If sUserID = "" Then Exit Function
    
    If spgpDetachedSigCreate(App.Path & "\PKI.txt", App.Path & "\PKI.pgp", sUserID, sPassword, _
        sComment, PGPHashAlgorithm_Default, 1) = 0 Then
        Dim ofso As FileSystemObject
        Set ofso = New FileSystemObject
        Dim otx As TextStream
        Set otx = ofso.OpenTextFile(App.Path & "\pki.pgp")
        ICHCSII_PKI_Sign = otx.ReadAll
        ofso.DeleteFile (App.Path & "\PKI.txt")
        ofso.DeleteFile (App.Path & "\PKI.pgp")
    End If

ErrHandler:


End Function

Private Function ICHCSII_PKI_Validate(ByVal SignedText As String, _
                                      ByVal Signature As String, _
                                            oUser As CHCSII_ClientSecurity.ICHCSII_User, _
                                            sigProp As SignatureProperties) As ValidationCode
On Error GoTo ErrHandler
    Dim bufferin As String
    Dim bufferOut As String
    Dim bufferoutLen As Long
    Dim SigProps As String
    Dim Sig() As String
    bufferoutLen = 3000
    bufferOut = Space(bufferoutLen)
    SigProps = Space(3000)
    Dim Verified As ValidationCode
    
'__ add the oUser keys to keyring
    If Not spgpKeyIsOnRing(oUser.GetAttribute("User_Ncid")) Then
        Call ImportKey(oUser)
    End If

'__ write out the two text strings to temp files

    Dim ofso As FileSystemObject
    Set ofso = New FileSystemObject
    Dim otx As TextStream
    
    Set otx = ofso.OpenTextFile(App.Path & "\pki.txt")
    Call otx.Write(SignedText)
    otx.Close
    
    Set otx = ofso.OpenTextFile(App.Path & "\pki.pgp")
    Call otx.Write(Signature)
    otx.Close
    
    Verified = spgpDetachedSigVerify(App.Path & "\PKI.pgp", App.Path & "\PKI.txt", SigProps)
    
    If Verified = SIGNED_GOOD Then
        Sig = Split(SigProps, vbTab)
        If (Sig(5) = 1 And Sig(6) = 1) Then
            
            If Sig(2) = oUser.GetAttribute("User_Ncid") Then
                ICHCSII_PKI_Validate = SIGNED_GOOD
            Else
                ICHCSII_PKI_Validate = SIGNED_BAD
            End If
        End If
    Else
        ICHCSII_PKI_Validate = Verified
    End If
    Exit Function
    
ErrHandler:
    ICHCSII_PKI_Validate = SIGNED_NO_KEY
End Function

Private Function GetUserId(oUser As CHCSII_ClientSecurity.ICHCSII_User) As String
    
End Function

Private Function ImportKey(oUser As CHCSII_ClientSecurity.ICHCSII_User) As String
On Error GoTo ErrHandler
    Dim bufferin As String
    Dim KeyProps As String
    Dim KeyPropsLen As Long
    KeyPropsLen = 2000
    KeyProps = Space(KeyPropsLen)
    goSecSQL.Execute "select keypair from user_pki where user_ncid = " & oUser.GetAttribute("User_NCID")
      
    If Not goSecSQL.EOF Then
        If Len(goSecSQL("keypair")) > 10 Then
            Debug.Print spgpKeyImport(goSecSQL("keypair"), KeyProps, KeyPropsLen, 1, 1)
        End If
    End If
ErrHandler:
End Function
