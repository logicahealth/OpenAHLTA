VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriorityType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarTypeCode As hlLABPRIORITY 'local copy
Private mvarText As String 'local copy
Private mvarHL7Code As String

Friend Property Let Text(ByVal vData As String)
    mvarText = vData
End Property

Public Property Get Text() As String
    Text = mvarText
End Property

Friend Property Let TypeCode(ByVal vData As hlLABPRIORITY)
    mvarTypeCode = vData
End Property

Public Property Get TypeCode() As hlLABPRIORITY
Attribute TypeCode.VB_UserMemId = 0
    TypeCode = mvarTypeCode
End Property

Friend Property Let HL7Code(ByVal szData As String)
     mvarHL7Code = szData
End Property

Public Property Get HL7Code() As String
     HL7Code = mvarHL7Code
End Property

