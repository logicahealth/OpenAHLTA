VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "privilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public PrivName As String
Public PrivDesc As String
Public PrivID   As Integer
Private msPrivString As String
Public pRead    As Boolean
Public pWrite   As Boolean
Public pDelete  As Boolean
Public pBreakGlass As Boolean
Public pNeedCosign As Boolean
Public pCanCosign As Boolean
Public pSign     As Boolean
Public AddRemove As String  '_ A or R

Public bChanged  As Boolean
Public bDeleted  As Boolean
Public bRedundant As Boolean   '- during ss import, this priv is already there

'<<
Public Property Let PrivString(RHV As String)
        pWrite = (InStr(RHV, "W") > 0)
        pRead = (InStr(RHV, "R") > 0)
        pDelete = (InStr(RHV, "D") > 0)
        pBreakGlass = (InStr(RHV, "B") > 0)
        pCanCosign = (InStr(RHV, "C") > 0)
        pNeedCosign = (InStr(RHV, "N") > 0)
        pSign = (InStr(RHV, "S") > 0)
        msPrivString = RHV
End Property

Public Property Get PrivString() As String
    Dim s As String
    s = ""
    If pRead Then s = s & "R"
    If pWrite Then s = s & "W"
    If pDelete Then s = s & "D"
    If pSign Then s = s & "S"
    If pCanCosign Then s = s & "C"
    If pBreakGlass Then s = s & "B"
    If pNeedCosign Then s = s & "N"
    
    msPrivString = s
    PrivString = s
End Property

Public Function DisplayPriv() As String
    Dim txt As String
    txt = "RWDSCBN"
    If Not pWrite Then txt = Replace(txt, "W", "-")
    If Not pRead Then txt = Replace(txt, "R", "-")
    If Not pDelete Then txt = Replace(txt, "D", "-")
    If Not pBreakGlass Then txt = Replace(txt, "B", "-")
    If Not pCanCosign Then txt = Replace(txt, "C", "-")
    If Not pNeedCosign Then txt = Replace(txt, "N", "-")
    If Not pSign Then txt = Replace(txt, "S", "-")
    DisplayPriv = txt
    
End Function

Public Function SamePriv(ByVal sTxt As String) As Boolean
    '-- see if incoming privstring has same          as mine
    Dim oP As New privilege
    oP.PrivString = sTxt
    If oP.pBreakGlass = pBreakGlass _
        And oP.pCanCosign = pCanCosign _
        And oP.pDelete = pDelete _
        And oP.pNeedCosign = pNeedCosign _
        And oP.pRead = pRead _
        And oP.pSign = pSign _
        And oP.pWrite = pWrite Then
        
        SamePriv = True
    Else
        SamePriv = False
    End If

End Function
