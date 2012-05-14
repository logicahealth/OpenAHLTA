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
Attribute VB_Name = "GUIDFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGUID As GUID) As Long
   
Private Const S_OK = 0 ' return value from CoCreateGuid
Private Const FORMAT_GUID = "@@@@@@@@-@@@@-@@@@-@@@@-@@@@@@@@@@@@"
Private Function GetGUID() As String
   
  Dim HRESULT As Long
  Dim pGUID As GUID
  Dim MyGuidString As String
  Dim MyGuidString1 As String
  Dim MyGuidString2 As String
  Dim MyGuidString3 As String
  Dim DataLen As Integer
  Dim StringLen As Integer
  Dim iLoop As Integer
  
  On Error GoTo error_olemsg
  
  HRESULT = CoCreateGuid(pGUID)
  
  If HRESULT = S_OK Then
  
     MyGuidString1 = Hex$(pGUID.Data1)
     StringLen = Len(MyGuidString1)
     DataLen = Len(pGUID.Data1)
     MyGuidString1 = LeadingZeros(2 * DataLen, StringLen) _
        & MyGuidString1 'First 4 bytes (8 hex digits)
  
     MyGuidString2 = Hex$(pGUID.Data2)
     StringLen = Len(MyGuidString2)
     DataLen = Len(pGUID.Data2)
     MyGuidString2 = LeadingZeros(2 * DataLen, StringLen) _
        & Trim$(MyGuidString2) 'Next 2 bytes (4 hex digits)
  
     MyGuidString3 = Hex$(pGUID.Data3)
     StringLen = Len(MyGuidString3)
     DataLen = Len(pGUID.Data3)
     MyGuidString3 = LeadingZeros(2 * DataLen, StringLen) _
        & Trim$(MyGuidString3) 'Next 2 bytes (4 hex digits)
  
     GetGUID = _
        MyGuidString1 & MyGuidString2 & MyGuidString3
  
     For iLoop = 0 To 7
        MyGuidString = MyGuidString & _
             Format$(Hex$(pGUID.Data4(iLoop)), "00")
     Next iLoop%
  
     'MyGuidString contains last 8 bytes of Guid (16 hex digits)
     GetGUID = GetGUID & MyGuidString
  
  Else
     GetGUID = "00000000" ' return zeros if function unsuccessful
  End If
  
  Exit Function
   
error_olemsg:
  MsgBox "Application Error: " & Err.Description & " Occured in DMBE.GUIDFactory.GetGUID" & "Error " & Str(Err) & ": " & Error$(Err)
  GetGUID = "00000000"
  Exit Function
   
End Function
Private Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
   
   LeadingZeros = String$(ExpectedLen - ActualLen, "0")

End Function


Public Function CreateGUID() As String
Attribute CreateGUID.VB_Description = "Returns the formatted string representation of the newly created GUID"
Attribute CreateGUID.VB_UserMemId = 0
    
    CreateGUID = Format(GetGUID(), FORMAT_GUID)

End Function

