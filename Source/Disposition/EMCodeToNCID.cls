VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMCodeToNCID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Object used to store EM Code and NCID in cache to minimize database calls
'SCR-10756

Private msEMCode As String
Private msEMNCID As String

Public Property Get EM_Code() As String
    
    EM_Code = msEMCode

End Property

Public Property Let EM_Code(ByVal sValue As String)

    msEMCode = sValue

End Property

Public Property Get EM_NCID() As String

    EM_NCID = msEMNCID

End Property

Public Property Let EM_NCID(ByVal sValue As String)

    msEMNCID = sValue

End Property
