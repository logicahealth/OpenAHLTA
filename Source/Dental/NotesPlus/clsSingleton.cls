VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSingleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Caries_Risk() As clsCariesRisk
    If gobjCariesRisk Is Nothing Then
        Set gobjCariesRisk = New clsCariesRisk
    End If
    
    Set Caries_Risk = gobjCariesRisk
End Property

'*********************************************************************
'Sub:Destroy_Objects
'Author:KDunne
'Date:3/12/2002
'Purpose: Destroys Singleton objects
'*********************************************************************
Public Sub Destroy_Objects()
    'SF 3-22-02, check to see if it is nothing first.
    If Not gobjCariesRisk Is Nothing Then
        gobjCariesRisk.Destroy_Objects
    End If
    
    Set gobjCariesRisk = Nothing
End Sub

