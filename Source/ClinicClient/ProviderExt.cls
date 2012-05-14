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
Attribute VB_Name = "mclsProviderExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           mclsProviderExt/ProviderExt.bas
'''
'''Description:      Provider record with extended set of properties
'''                  including SSN and Phone numbers
'''<SCR  Pharmacy 301
'''
'''Modules Required:
'''
'''
'''Note:
'''
'''Author         :  Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME      As String = "mclsProviderExt."


''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''
'''Properties
Private msNCID                   As String
Private msNameFirst              As String
Private msNameLast               As String
Private msNameMiddle             As String
Private msNameFull               As String
Private msPhone1                 As String
Private msPhone2                 As String
Private msProviderIEN            As String
Private msSignatureClass         As String
Private msUserIEN                As String
Private msSSN                    As String
Private msSpecialityCode         As String
Private msDefaultClinicNCID      As String
Private msFacilityNCID           As String
Friend Function F_bDeserialize(objFields As adodb.Fields) As Boolean
'''************************************************************************
'''Called By:     ClinicClient_objProviderExt
'''
'''Arguments:
''' objFields
'''   Provider Record requested to be used as source
'''
'''Description:   Extract provider properties from specified recordt
'''               and populate object properties
'''
'''Note:          This procedure expects record from provider table
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "F_bDeserialize"

On Error Resume Next

200   Call M_bDecodeName(objFields("NAME").Value & "")

220   msNCID = objFields("NCID").Value & ""
260   msPhone1 = objFields("DUTY_PHONE_1").Value & ""
280   msPhone2 = objFields("DUTY_PHONE_2").Value & ""
300   msProviderIEN = objFields("IEN").Value & ""
320   msSignatureClass = objFields("SIGNATURE_CLASS").Value & ""
340   msUserIEN = objFields("USER_IEN").Value & ""
380   msSSN = objFields("SSN").Value & ""
400   msSpecialityCode = objFields("SPECIALTY_CODE").Value & ""
420   msDefaultClinicNCID = objFields("CLINIC_NCID").Value & ""
440   msFacilityNCID = objFields("FACILITY_NCID").Value & ""
    
999   F_bDeserialize = True
    
End Function

Private Function M_bDecodeName(ByVal sFullName As String) As Boolean
'''************************************************************************
'''Called By:     F_bDeserialize
'''
'''Arguments:
''' sFullName
'''   Full name
'''
'''Description:   Extract from specified full name first and last names
'''
'''Note:          Full name is expected to be in the following format:
'''               LastName,FirstName Middlename
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bDecodeName"

On Error GoTo ErrHnd

20       msNameFull = sFullName
100      msNameFirst = LibStr_sExtractSubstring(sFullName, ",", " ")
200      Call LibStr_sExtractSubstring(sFullName, msNameLast, msNameMiddle, "," & msNameFirst)
300      msNameFirst = Trim$(msNameFirst)
320      msNameLast = Trim$(msNameLast)
340      msNameMiddle = Trim$(msNameMiddle)

999   M_bDecodeName = True
   
Exit Function
ErrHnd:
   M_bDecodeName = False
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Friend Property Get fobjClone() As mclsProviderExt
'''************************************************************************
'''Called By:     ClinicClient_objProviderExt
'''
'''Arguments:
'''
'''Description:   Create and return copy of it self
'''
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Returns:       If successful
'''                  Copy of the this object
'''               Else
'''                  nothing
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "fobjClone"

   Dim objResult        As mclsProviderExt
   
On Error GoTo ErrHnd

20    Set objResult = New mclsProviderExt
100   With objResult
200      .fsNCID = msNCID
210      .fsNameFirst = msNameFirst
220      .fsNameLast = msNameLast
230      .fsNameMiddle = msNameMiddle
232      .fsNameFull = msNameFull
240      .fsPhone1 = msPhone1
250      .fsPhone2 = msPhone2
260      .fsProviderIEN = msProviderIEN
270      .fsSignatureClass = msSignatureClass
280      .fsUserIEN = msUserIEN
290      .fsSSN = msSSN
300      .fsSpecialityCode = msSpecialityCode
310      .fsDefaultClinicNCID = msDefaultClinicNCID
320      .fsFacilityNCID = msFacilityNCID
      End With
    
999   Set fobjClone = objResult
   
Exit Property
ErrHnd:
   Set fobjClone = Nothing
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
'   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property
Friend Property Let fsNCID(ByVal sVal As String): msNCID = sVal: End Property
Friend Property Let fsNameFirst(ByVal sVal As String): msNameFirst = sVal: End Property
Friend Property Let fsNameLast(ByVal sVal As String): msNameLast = sVal: End Property
Friend Property Let fsPhone1(ByVal sVal As String): msPhone1 = sVal: End Property
Friend Property Let fsPhone2(ByVal sVal As String): msPhone2 = sVal: End Property
Friend Property Let fsProviderIEN(ByVal sVal As String): msProviderIEN = sVal: End Property
Friend Property Let fsSignatureClass(ByVal sVal As String): msSignatureClass = sVal: End Property
Friend Property Let fsUserIEN(ByVal sVal As String): msUserIEN = sVal: End Property
Friend Property Let fsSSN(ByVal sVal As String): msSSN = sVal: End Property
Friend Property Let fsSpecialityCode(ByVal sVal As String): msSpecialityCode = sVal: End Property
Friend Property Let fsDefaultClinicNCID(ByVal sVal As String): msDefaultClinicNCID = sVal: End Property
Friend Property Let fsFacilityNCID(ByVal sVal As String): msFacilityNCID = sVal: End Property
Friend Property Let fsNameFull(ByVal sVal As String): msNameFull = sVal: End Property
Public Property Get psNCID() As String: psNCID = msNCID: End Property
Public Property Get psNameFirst() As String: psNameFirst = msNameFirst: End Property
Public Property Get psNameLast() As String: psNameLast = msNameLast: End Property
Public Property Get psPhone1() As String: psPhone1 = msPhone1: End Property
Public Property Get psPhone2() As String: psPhone2 = msPhone2: End Property
Public Property Get psProviderIEN() As String: psProviderIEN = msProviderIEN: End Property
Public Property Get psSignatureClass() As String: psSignatureClass = msSignatureClass: End Property
Public Property Get psUserIEN() As String: psUserIEN = msUserIEN: End Property
Public Property Get psSSN() As String: psSSN = msSSN: End Property
Public Property Get psSpecialityCode() As String: psSpecialityCode = msSpecialityCode: End Property
Public Property Get psDefaultClinicNCID() As String: psDefaultClinicNCID = msDefaultClinicNCID: End Property
Public Property Get psFacilityNCID() As String: psFacilityNCID = msFacilityNCID: End Property
Friend Property Let fsNameMiddle(ByVal sVal As String): msNameMiddle = sVal: End Property
Public Property Get psNameMiddle() As String: psNameMiddle = msNameMiddle: End Property
Public Property Get psNameFull() As String: psNameFull = msNameFull: End Property
