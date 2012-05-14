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
Attribute VB_Name = "mclsLibUSPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:              mclsLibUSPhone/LibUSPhone.cls
'''
'''Description:         Allows to process, evaluate and store an entry as a US phone number
'''                     Input
'''                        Formated or non-formated string representations or Number
'''                        as a single entry: psPhone
'''                        or as sub-fields: psAreaCode, psExchange, psLine, psExtension
'''                     Output
'''                        psPhone - Phone number without formated characters
'''                        psPhoneFormated - Phone number formated according to DOD style guide
'''                        Sub-codes: psAreaCode, psExchange, psLine, psExtension
'''
'''Usage:
'''Modules Required:
'''
'''Note:                Area code, Exchange and Line number are required sub fields - Extension optional
'''                     Input may be prefaced with "1-" or "1" but this part of the entry
'''                     will not be persisted in the output
'''Warning:
'''
'''Author         :     Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME            As String = "mclsLibUSPhone."
Private Const mlSTANDARD_PHONE_LEN     As Long = 10   '''Area Code + Exchange + Line

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private msPhoneInitial        As String            '''Original value
Private msCountryCode         As String
Private msAreaCode            As String
Private msExchange            As String
Private msLine                As String
Private msExtension           As String
Private mlMaxExtension        As Long
Private mlMinExtension        As Long
Private mbAllowAlpha          As Boolean
Private mlExtensionRequired   As pnumYNA_YesNoAny

''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

Public Event pevChanged(ByVal sNewValue As String)
Public Property Get pbAllowAlpha() As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current property settings
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Current property settings
'''                  Else
'''                     false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plExtensionRequired"
   
On Error GoTo ErrHnd

   pbAllowAlpha = mbAllowAlpha
   
Exit Property
ErrHnd:
   pbAllowAlpha = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property


Public Property Let pbAllowAlpha(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' lVal
'''   New property setting
'''
'''Description:         R W Default
'''                     Y Y False
'''                     Indicates whether letters a-y is valid input characters.
'''                     If true - these letters will be converted to corresponding digits on input
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pbAllowAlpha"
   
On Error GoTo ErrHnd

   mbAllowAlpha = bVal
   If mbAllowAlpha Then
      psPhone = msPhoneInitial '''Reprocess original value to make sure all letters are converted to digits
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property


Public Property Let plExtensionRequired(ByVal lVal As pnumYNA_YesNoAny)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' lVal
'''   New property setting
'''
'''Description:         R W Default
'''                     Y Y Any
'''                     Indicates whether extension is required - YES,
'''                     not allowed - NO or optional - ANY
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plExtensionRequired"
   
On Error GoTo ErrHnd

   mlExtensionRequired = lVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Get plExtensionRequired() As pnumYNA_YesNoAny
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current property settings
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Current property settings
'''                  Else
'''                     -1
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plExtensionRequired"
   
On Error GoTo ErrHnd

   plExtensionRequired = mlExtensionRequired
   
Exit Property
ErrHnd:
   plExtensionRequired = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Let plMaxExtension(ByVal lVal As Long)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' lVal
'''   New property value
'''
'''Description:      R W Default
'''                  Y Y 4
'''                  Indicates maximum allowable length of the phone extension
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plMaxExtension"
   
On Error GoTo ErrHnd

   mlMaxExtension = lVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Get plMaxExtension() As Long
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current property settings
'''
'''Updates Globals:
'''Returns:          If successful
'''                     current property settings
'''                  Else
'''                     -1
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plMaxExtension"
   
On Error GoTo ErrHnd
   
   plMaxExtension = mlMaxExtension
   
Exit Property
ErrHnd:
   plMaxExtension = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Let psAreaCode(ByVal sVal As String)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sVal
'''   New Area Code
'''
'''Description:      R W Default
'''                  Y Y Empty string
'''                  Sets new area code
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let psAreaCode"
   
On Error GoTo ErrHnd
   
   If pbAllowAlpha Then
     Call P_bDecodeAlphaToNum(sVal, sVal)
   End If
   msAreaCode = sVal
   RaiseEvent pevChanged(psPhoneFormated)
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Get psAreaCode() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current Area Code
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Area Code
'''                  Else
'''                     Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psAreaCode"
   
On Error GoTo ErrHnd

   psAreaCode = msAreaCode
   
Exit Property
ErrHnd:
   psAreaCode = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Let psExchange(ByVal sVal As String)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sVal
'''   New Exchange Code
'''
'''Description:      R W Default
'''                  Y Y Empty string
'''                  Sets new New Exchange Code
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let psExchange"
   
On Error GoTo ErrHnd

   If pbAllowAlpha Then
     Call P_bDecodeAlphaToNum(sVal, sVal)
   End If
   msExchange = sVal
   RaiseEvent pevChanged(psPhoneFormated)
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get psExchange() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current Exchange Code
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Current Exchange Code
'''                  Else
'''                     Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psExchange"
   
On Error GoTo ErrHnd
   
   psExchange = msExchange
   
Exit Property
ErrHnd:
   psExchange = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Let psLine(ByVal sVal As String)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sVal
'''   New value for phone line property
'''
'''Description:      R W Default
'''                  Y Y Empty string
'''                  New value for phone line property
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let psLine"
   
On Error GoTo ErrHnd

   If pbAllowAlpha Then
     Call P_bDecodeAlphaToNum(sVal, sVal)
   End If
   msLine = sVal
   RaiseEvent pevChanged(psPhoneFormated)
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Get psLine() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current Phone Line Number
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Line Number
'''                  Else
'''                     Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psLine"
   
On Error GoTo ErrHnd

   psLine = msLine
   
Exit Property
ErrHnd:
   psLine = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Let psExtension(ByVal sVal As String)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sVal
'''   New property value
'''
'''Description:      R W Default
'''                  Y Y Empty string
'''                  Sets value of th ephone extension property
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let psExtension"
   
On Error GoTo ErrHnd

   If pbAllowAlpha Then
     Call P_bDecodeAlphaToNum(sVal, sVal)
   End If
   msExtension = sVal
   RaiseEvent pevChanged(psPhoneFormated)
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get psExtension() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current phone Extension
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Curent phone Extension
'''                  Else
'''                     Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psExtension"
   
On Error GoTo ErrHnd

   psExtension = msExtension
   
Exit Property
ErrHnd:
   psExtension = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property
Public Property Get plMaxInputLen() As Long
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      R W Default
'''                  Y N N/A
'''                  Returns current maximum length for Phone number entry
'''                  Based on plExtensionRequired and plMaxExtension
'''                  if extension is not allowed (plExtensionRequired=lYNA_No)
'''                     Max input length is 15
'''                     i.e. len("1-(999)999-9999")
'''                  else (plExtensionRequired = lYNA_Yes or plExtensionRequired = lYNA_ANY)
'''                     Max input length is 16 + plMaxExtension
'''                     i.e. len("1-(999)999-9999x" & plMaxExtension)
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Current maximum allowable length of input for phone number property
'''                  Else
'''                     -1
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plMaxInputLen"
   
   Dim lResult          As Long
   
On Error GoTo ErrHnd
   
20    If plExtensionRequired = lYNA_No Then
40       lResult = 15
      Else
100      lResult = 16 + plMaxExtension
      End If
   
999   plMaxInputLen = lResult
   
Exit Property
ErrHnd:
   plMaxInputLen = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Function P_bDecodeAlphaToNum(ByVal sPhoneIn As String, ByRef sPhoneOut As String) As Boolean
'''************************************************************************
'''Called By:     Let psPhone
'''               Let psAreaCode
'''               Let psExchange
'''               Let psLine
'''               Let psExtension
'''Arguments:
''' sPhoneIn
'''   Phone number requested to be decodes
''' sPhoneOut
'''   Result. Returned by ref
'''
'''Description:      Replace all leters of the specified phone number with corresponding digits
'''
'''Note:
'''
'''Updates Globals:
'''Returns:          If succcessful
'''                     true
'''                  else
'''                     false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bDecodeAlphaToNum"
   
   Dim lLen             As Long
   Dim lPos             As Long

On Error GoTo ErrHnd
   
   sPhoneOut = sPhoneIn
   lLen = Len(sPhoneOut)
   For lPos = 1 To lLen
      Select Case UCase$(Mid$(sPhoneOut, lPos, 1))
         Case "A", "B", "C"
            Mid$(sPhoneOut, lPos, 1) = "2"
         Case "D", "E", "F"
            Mid$(sPhoneOut, lPos, 1) = "3"
         Case "G", "H", "I"
            Mid$(sPhoneOut, lPos, 1) = "4"
         Case "J", "K", "L"
            Mid$(sPhoneOut, lPos, 1) = "5"
         Case "M", "N", "O"
            Mid$(sPhoneOut, lPos, 1) = "6"
         Case "P", "R", "S"
            Mid$(sPhoneOut, lPos, 1) = "7"
         Case "T", "U", "V"
            Mid$(sPhoneOut, lPos, 1) = "8"
         Case "W", "X", "Y"
            Mid$(sPhoneOut, lPos, 1) = "9"
         Case Else
            '''Do nothing - keep as is
      End Select
   Next lPos
   
   P_bDecodeAlphaToNum = True

Exit Function
ErrHnd:
   P_bDecodeAlphaToNum = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Public Property Get pbValid _
      (Optional ByRef sErrMsg As String, _
       Optional ByVal vPhone As Variant, _
       Optional ByRef sPhoneResult As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sErrMsg
'''   Error info if invalid - returned by ref
''' vPhone
'''   Optional value requested to be evaluated. If not specified current phone number will be evaluated
''' sPhoneResult
'''   Optional output - phone number without formated characters. returned byval
'''
'''Description:      R W Default
'''                  Y N N/A
'''                  If specified - evaluate specified phone number
'''                  Else - evaluate current phone number
'''
'''Note:
'''
'''Updates Globals:
'''Returns:          True  - If requested phone number is valid
'''                  Else  - False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pbValid"
   
   Dim bErr             As Boolean
   Dim lLen             As Long
   
On Error GoTo ErrHnd

20    sErrMsg = ""
      '''If not specified
100   If IsMissing(vPhone) Then
120      sPhoneResult = psPhone  '''Use current phone number associated with this object
      Else
200      sPhoneResult = gobjLibValid.LibValid_vCastType(vPhone)     '''Use specified entry
         '''If entry is not to long
         If Len(sPhoneResult) <= plMaxInputLen Then
            '''Remove dial one if any(e.g. 1(999) .... or 1-(999) .... or).
            '''Note: Dash if any will be remove as part of format removal bellow
220         sPhoneResult = gobjLibStr.LibStr_sStartWith(sPhoneResult, "1", False, True)
300         If pbAllowAlpha Then
320            bErr = Not P_bDecodeAlphaToNum(sPhoneResult, sPhoneResult)
            End If
         Else
            bErr = True
            sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsVALUE, gsIS, gsTO, gsLONG)
         End If
      End If
   
400   If Not bErr Then
         '''Remove format
420      sPhoneResult = gobjLibStr.LibStr_sReplaceChar(sPhoneResult, gobjLibStr.psNUMBERS, bErr, True)
500      If Not bErr Then
            '''At this point entry should contain numbers only
520         lLen = Len(sPhoneResult)
            '''If number of digits is sufficient to constitute valid phone number
            Select Case True
               Case lLen = mlSTANDARD_PHONE_LEN    '''Extension is not specified
                  '''If extension is required
600               If plExtensionRequired = lYNA_Yes Then
620                  sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsMISSING, gsPHONE, gsEXTENSION)
                  End If
               Case lLen > mlSTANDARD_PHONE_LEN    '''Extension is specified
                  '''If extension is not allowed
700               If (plExtensionRequired = lYNA_No) Then
720                  sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPHONE, gsEXTENSION, gsIS, gsNOT, gsALLOWED)
                  Else
760                  Call gobjLibValid.LibValid_bInteger _
                        (lLen - mlSTANDARD_PHONE_LEN, plMinExtension, plMaxExtension, sErrMsg, _
                         gobjLibStr.LibStr_sBuildSentence(gsLENGTH, gsOF, gsPHONE, gsEXTENSION) _
                        )
                  End If
               Case Else   '''less digits then mlSTANDARD_PHONE_LEN
800               sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsINSUFFICIENT, gsENTRY)
            End Select
         End If
      End If

900   If bErr And Len(sErrMsg) = 0 Then
920      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsERROR, gsPROCESSING, gsINPUT, "-", Erl)
      End If
      
999   pbValid = Len(sErrMsg) = 0

Exit Property
ErrHnd:
   sErrMsg = Err.Description
   pbValid = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Property Get psPhone() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      R W Default
'''                  Y Y Empty string
'''                  Returns current phone number
'''
'''Note:             This property will always return non-empty Phone Extension
'''                  regardlees whether its required or not vs. psPhoneFormated
'''                  which will only return it if extension is required or optional
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Current non-formated Phone Number
'''                  Else
'''                     Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psPhone"
   
   Dim sResult          As String
   
On Error GoTo ErrHnd

20    sResult = msAreaCode & msExchange & msLine & msExtension
   
999   psPhone = sResult

Exit Property
ErrHnd:
   psPhone = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Let psPhone(ByVal sVal As String)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sVal
'''   New complete phone number
'''
'''Description:      R W Default
'''                  Y Y Empty string
'''                  Accept specified phone number - update corresponding sub properties
'''
'''Note:
'''Warning:          Specified value will be accepted regardless of whether it
'''                  represent valid phone number or not
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "Let psPhone"
   
On Error GoTo ErrHnd

20    msPhoneInitial = sVal   '''store original input value
      '''Process specified phone number to make sure it contain digits only
40    If pbValid(, sVal, sVal) Then
         '''Do nothing - Accept any value
      End If
      '''Extract and store sub-fields
100   msAreaCode = Left$(sVal, 3)
200   msExchange = Mid$(sVal, 4, 3)
300   msLine = Mid$(sVal, 7, 4)
400   msExtension = Mid$(sVal, 11)

900   RaiseEvent pevChanged(psPhoneFormated)

      
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Get psPhoneFormated() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      R W Default
'''                  Y N <Area Code>-<Exchange Code>-<Line Number>[x<Extension>]
'''                  Format and return current phone number
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Current phone number formated according to DOD style guide
'''                  Else
'''                     Empty string
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psPhoneFormated"
   
   Dim sResult          As String
   
On Error GoTo ErrHnd


20    sResult = msAreaCode & "-" & msExchange & "-" & msLine

      '''If extension is required or (allowed and specified)
100   If mlExtensionRequired = lYNA_Yes Or _
        (mlExtensionRequired = lYNA_Any And Len(msExtension) > 0) _
      Then
200      sResult = sResult & "x" & msExtension
      End If
   
999   psPhoneFormated = sResult
   
Exit Property
ErrHnd:
   psPhoneFormated = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Property

Public Property Get psPhoneInitial() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      R W Default
'''                  Y N empty String
'''                  Return original input value
'''
'''Updates Globals:
'''Returns:          If successful
'''                     Return last input value of the psPhone property
'''                  Else
'''                     Empty string
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psPhoneInitial"
   
   
On Error GoTo ErrHnd

   psPhoneInitial = msPhoneInitial
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Property

Private Sub Class_Initialize()
   mlMaxExtension = 4
   mlMinExtension = 1
End Sub

Public Property Let plMinExtension(ByVal lVal As Long)
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' lVal
'''   New property value
'''
'''Description:      R W Default
'''                  Y Y 1
'''                  Indicates minimum allowable length of the phone extension
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plMinExtension"
   
On Error GoTo ErrHnd

   mlMinExtension = lVal
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property
Public Property Get plMinExtension() As Long
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:      Returns current property settings
'''
'''Updates Globals:
'''Returns:          If successful
'''                     current property settings
'''                  Else
'''                     -1
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get plMinExtension"
   
On Error GoTo ErrHnd
   
   plMinExtension = mlMinExtension
   
Exit Property
ErrHnd:
   plMinExtension = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

