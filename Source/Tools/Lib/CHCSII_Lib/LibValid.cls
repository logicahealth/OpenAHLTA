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
Attribute VB_Name = "gclsLibValid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibValid/LibValid.cls
'''
'''Description:   Validation, Type casting, Rounding  utilities and associated entities.
'''
'''Modules Required:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "gclsLibValid."

'''Data types supported by this module
Public Enum gnumDTYPDataType
   lDTYP_String = vbString
   lDTYP_Integer = vbInteger
   lDTYP_Long = vbLong
   lDTYP_Double = vbDouble
   lDTYP_Date = vbDate
   lDTYP_Boolean = vbBoolean
   lDTYP_Time = -990
   lDTYP_Path = -991
   lDTYP_NetworkPath = -992
   lDTYP_DateTime = -993
   lDTYP_USPhone = -994
   lDTYP_InterPhone = -995
End Enum

Public pbLibValidMsgPop    As Boolean
Public pbLibValidMsgBeep   As Boolean

Private mobjUSPhone        As mclsLibUSPhone
Function LibValid_bIsNum(ByVal vNum As Variant) As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' vNum
'''   Value requested to be evaluated
'''
'''Description:   Evaluate whether specified value represents a number
'''
'''Note:    See also LibValid_dVal
'''
'''Warning: Numbers expressed in exponential notation e.g. 1E1 or 9D1 are evaluated as not numeric
''          Empty is evaluated as not numeric
'''
'''Updates Globals:
'''
'''Returns:          If specifed string represent a number
'''                     true
'''                  Else
'''                     false
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bIsNum"

   Dim bResult    As Boolean
   
On Error GoTo ErrHnd
   
100   If IsNumeric(vNum) Then
200      If Not IsEmpty(vNum) Then
220         bResult = gobjLibStr.LibStr_bValidString(vNum, "ED", False, True)
         End If
      End If
   
999   LibValid_bIsNum = bResult
   
Exit Function
ErrHnd:
   LibValid_bIsNum = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)
   
End Function

Function LibValid_dVal(ByVal sStr As String) As Double
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sStr
'''   String expression requested to be converted to and returned as number
'''
'''Description:   If specified string starts with number convert to Dbl and return it
'''               else return zero
'''
'''Note:          Input/Output differences between This function(LibValid_dVal) and Val() function - its VB couterpart
'''               Numbers expressed in exponential notation:
'''                  Input: "1E1"      / Output: this function 1        - Val function 10
'''                  Input: "9.2D2"    / Output: this function 9.2      - Val function 920
'''
'''               See also LibValid_bIsNum, LibValid_dValExt
'''Warning:
'''Updates Globals:
'''
'''Returns:          If specifed string starts with number
'''                     this number
'''                  Else
'''                     Zero
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME                 As String = "LibValid_dVal"

   Dim dResult                      As Double
   Dim sResult                      As String
   Dim lPos                         As Long
   Dim bOneDot                      As Boolean
   Dim sChar                        As String
   
On Error GoTo ErrHnd

20    sStr = Trim$(sStr)
100   For lPos = 1 To Len(sStr)
120      sChar = Mid$(sStr, lPos, 1)
200      If InStr("0123456789", sChar) > 0 Then
220         sResult = sResult & sChar
300      ElseIf sChar = "." Then
400         If Not bOneDot Then
420            bOneDot = True
440            sResult = sResult & sChar
            Else
500            Exit For
            End If
         Else
800         Exit For
         End If
      Next lPos

900   If Len(sResult) > 0 Then
920      dResult = Val(sResult)
      End If
      
999   LibValid_dVal = dResult
   
Exit Function
ErrHnd:
   LibValid_dVal = 0
   Call gobjLibApp.LibApp_Error _
         (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Function LibValid_bBoolean _
(ByVal vValue As Variant, Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "") _
 As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments:
''' vValue        Value to be evaluated as boolean
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   Check whether specified value can be evaluated as boolean
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bBoolean"

   Dim sVal             As String
   
On Error GoTo ErrHnd
  
   
   If sCaption = "" Then
      sCaption = gsVALUE
   End If
   
   Call LibValid_bEvaluateBoolean(vValue, sErrMsg)
   
   If sErrMsg <> "" Then
      sErrMsg = sCaption & " - " & sErrMsg
   End If

   LibValid_bBoolean = sErrMsg = ""
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bBoolean = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")

End Function
Public Function LibValid_bEvaluateBoolean _
               (ByVal vValue As Variant, Optional ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' vValue        Value to be evaluated as date
''' sErrMsg       Used to return error information if value is invalid
'''
'''Description:   Check whether specified value can be evaluated as boolean
'''               If Yes return its boolean representation
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If no errors and specified value can be evaluated as boolean and evaluates to true
'''                  boolean representation of the specified value
'''               Else
'''                  False and sErrMsg
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bEvaluateBoolean"

   Dim sVal             As String
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
  
   sErrMsg = ""
   
   Select Case VarType(vValue)
      Case vbBoolean
         bResult = vValue
      Case vbLong, vbInteger, vbSingle, vbDouble
         Select Case vValue
            Case 1, -1
               bResult = True
            Case 0
            Case Else
               sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsAMBIGUOUS, gsVALUE)
         End Select
      Case vbString
         Select Case UCase(Trim$(vValue))
            Case "TRUE", "YES", "ON", "T", "Y", "1", "-1"
               bResult = True
            Case "FALSE", "NO", "OFF", "F", "N", "0"
            Case Else
               sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsUNDEFINED, gsTEXT, gsREPRESENTATION)
         End Select
      Case vbEmpty, vbNull
         bResult = False
      Case Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsINCOMPATIBLE, gsDATA, gsTYPE)
   End Select
   
   Call M_ProcessMsg(sErrMsg, "")

   LibValid_bEvaluateBoolean = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bEvaluateBoolean = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")

End Function
Public Function LibValid_bInterPhone _
               (ByVal vValue As Variant, Optional ByRef sErrMsg As String, _
                Optional ByVal sCaption As String = "Phone", Optional ByRef sNumber As String _
               ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments
''' vValue        Value to be evaluated as international phone number
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
''' sNumber       Optional Specified phone without formatting character. Rbyref
'''
'''Description:   Evaluate whether specified value is of the valid international phone number
'''
'''Note:          Due to vast variety of format for these entity evaluation criteria is fairly simple:
'''               Only allows digits and limited set of delimiters: (+()-x)
'''               Overall length should not exceed 25 characters(21 US: 1(999)-999-9999x9999 plus + and 3 digits country code ).
'''               Must have from 7 to 17 digits (14 US plus 3 digits country code)
'''               And it should not be all zeroes
'''
'''Warning:       Value will not be trimmed prior to validation
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bInterPhone"
   Const sVALID_DLM     As String = "+()-xX"
   
   Dim sPhone           As String
   Dim bErr             As Boolean
   
On Error GoTo ErrHnd
   
20    sErrMsg = ""
40    sNumber = ""
60    sPhone = gobjLibStr.LibStr_sEmptyIfNull(vValue, False)
      
      
100   With gobjLibStr
200      If LibValid_bInteger(Len(sPhone), 7, 25, sErrMsg, gsLENGTH) Then
            '''If entry consists of the valid characters only
300         If .LibStr_bValidString(sPhone, sVALID_DLM & .psNUMBERS, True) Then
               '''Strip of all non-digit characters
400            sPhone = .LibStr_sReplaceChar(sPhone, .psNUMBERS, bErr, True, "")
               If Not bErr Then
420               sNumber = sPhone
500               If LibValid_bInteger _
                     (Len(sPhone), 7, 17, sErrMsg, .LibStr_sBuildSentence(gsNUMBER, gsOF, gsDIGIT & "s")) _
                  Then
                     If LibValid_dVal(sPhone) = 0 Then
900                     sErrMsg = .LibStr_sBuildSentence(gsINVALID, gsNUMBER)
                     End If
                  End If
               Else
920               sErrMsg = .LibStr_sBuildSentence(gsPROCESSING, gsERROR)
               End If
            Else
940            sErrMsg = .LibStr_sBuildSentence(gsINVALID, gsFORMAT)
            End If
         End If
      End With
      
990   Call M_ProcessMsg(sErrMsg, sCaption)

999   LibValid_bInterPhone = sErrMsg = ""
   
Exit Function
ErrHnd:
   LibValid_bInterPhone = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sCaption & vbCrLf & gsLINE_ & Erl)

End Function

Public Function LibValid_bNumberRound _
       (ByRef vValue As Variant, _
       Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "", _
       Optional ByVal lNoOfDigitsAfterDecimal As Long = 2 _
      ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments
''' vValue        Value to be evaluated as number. Rounded if requested
''' vLow          minimum value allowed
''' vHigh         maximum value allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
''' lNoOfDigitsAfterDecimal  Maximum Number of digits after decimal.
'''               Specified number will be rounded to it. Ignored if less then zero
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:          This function will round specified number
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bNumberRound"

On Error GoTo ErrHnd
   
   Call M_bValidNumber(vValue, vLow, vHigh, sErrMsg, sCaption, , lNoOfDigitsAfterDecimal)

   LibValid_bNumberRound = sErrMsg = ""
   
Exit Function
ErrHnd:
   LibValid_bNumberRound = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sCaption)

End Function

Public Function LibValid_bUSPhone _
      (ByVal vValue As Variant, _
       Optional ByVal vMinExt As Variant = 0, Optional ByVal vMaxExt As Variant = 4, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "Phone", _
       Optional ByRef sFormated As String _
      ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments:
''' vValue        Value to be evaluated as US phone number
''' vMinExt       Minimum length of Phone Extension. Zero - Interpreted as Extension is optional.
'''               Greater then zero Interpreted as Extension is required
''' vMaxExt       Maximum Length of Phone Extension. Zero - Interpreted as Extension is not allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
''' sFormated     Optional specified phone number formated according to DOD style guide
'''               Returned byref
'''
'''Description:   Evaluate whether specified value is of the valid US phone number
'''
'''Note:          Area Code is required
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bUSPhone"
   
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd

   If mobjUSPhone Is Nothing Then
      Set mobjUSPhone = New mclsLibUSPhone
   End If
   With mobjUSPhone
      If M_bUseCriteria(vMinExt) Then
         If vMinExt > 0 Then
            .plExtensionRequired = lYNA_Yes
         Else
            .plExtensionRequired = lYNA_Any
         End If
         .plMinExtension = vMinExt
      End If
      If M_bUseCriteria(vMaxExt) Then
         If vMaxExt = 0 Then
            .plExtensionRequired = lYNA_No
         End If
         .plMaxExtension = vMaxExt
      End If
      .psPhone = vValue
      bResult = .pbValid(sErrMsg)
      sFormated = .psPhoneFormated
   End With
   
   Call M_ProcessMsg(sErrMsg, sCaption)
   LibValid_bUSPhone = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bUSPhone = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)
   
End Function

Public Function LibValid_bTypeIsSupported _
      (ByVal lType As gnumDTYPDataType, ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lType         Indicated data type
''' sErrMsg       Used to return error information if value is invalid
'''
'''Description:   Evaluate whether specified data type is supported
'''
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value data type is supported
'''                  True
'''               Else or error
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bTypeIsSupported"

On Error GoTo ErrHnd
  
   '''Default to valid
   sErrMsg = ""
   
   Select Case lType
      Case lDTYP_Long
      Case lDTYP_Double
      Case lDTYP_Date
      Case lDTYP_DateTime
      Case lDTYP_String
      Case lDTYP_Integer
      Case lDTYP_Path
      Case lDTYP_NetworkPath
      Case lDTYP_Boolean
      Case lDTYP_Time
      Case lDTYP_USPhone
      Case lDTYP_InterPhone
      Case Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsSPECIFIED, gsTYPE, gsIS, gsNOT, gsSUPPORTED)
   End Select
   
   LibValid_bTypeIsSupported = (sErrMsg = "")
   
Exit Function
ErrHnd:
   LibValid_bTypeIsSupported = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsTYPE & lType)
   
End Function
Public Function LibValid_bValue _
      (ByVal lType As gnumDTYPDataType, ByVal vValue As Variant, _
       Optional ByVal vMin As Variant, Optional ByVal vMax As Variant, Optional vValidValues As Variant, _
       Optional ByRef sErrMsg As String, _
       Optional ByVal sCaption As String = "" _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' vValue        Value to be evaluated
''' vMin          For string
'''                  Minimum number of characters
'''               For Date and numbers
'''                  Minimum value allowed
'''               For Path and NetworkPath
'''                  Not used
'''               For USPhone
'''                  Minimum length of the phone extension.
'''                  Zero indicates that Extension is optional.
'''                  Greater then zero interpreted as Extension is required
''' vMax          For string
'''                  Maximum number of characters
'''               For Date and numbers
'''                  Maximum value allowed
'''               For Path and NetworkPath
'''                  Not used
'''               For USPhone
'''                  Maximum length of the phone extension
'''                  Zero - indicates that Extension is not allowed
'''
''' vValidValues  Comma delimited list of valid values or one dimensional array
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "LibValid_bValue"

   Dim bResult             As Boolean
   Dim sValidDblValues()   As String
   Dim bListAsNumber       As Boolean
   
On Error GoTo ErrHnd
   
   bResult = M_bValidVal(lType, vValue, vMin, vMax, sErrMsg, sCaption)
   
   If bResult Then
      '''If valid values are requested to be evaluated and data type is not string(already evalauted)
      If M_bUseCriteria(vValidValues) Then
         If lType = lDTYP_Double And VarType(vValidValues) = vbString Then
            '''Convert to Array - if list contains unvalid numeric values will cause an error
            sValidDblValues = Split(vValidValues, ",")
            vValidValues = sValidDblValues
            bListAsNumber = True
         End If
         '''Note: caller may specify to have only particular Path(s) to be valid
         bResult = LibValid_bListedValue(vValue, vValidValues, , sCaption, bListAsNumber, sErrMsg)
      End If
   End If
   
   LibValid_bValue = bResult
   
Exit Function
ErrHnd:
   LibValid_bValue = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sCaption)
   
End Function
Public Function LibValid_bDateTime _
  (ByVal vValue As Variant, Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
   Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "" _
  ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments:
''' vValue        Value to be evaluated as date
''' vLow          Lowest Date-time allowed
''' vHigh         Highest Data-time allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bDateTime"
   
On Error GoTo ErrHnd
   
   LibValid_bDateTime = M_bValidDateTime(vValue, vLow, vHigh, sErrMsg, sCaption, False)
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bDateTime = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)
   
End Function
Public Function LibValid_bInteger _
      (ByVal vValue As Variant, Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "" _
      ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments
''' vValue        Value to be evaluated as integer
''' vLow          minimum value allowed
''' vHigh         maximum value allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
''' bFix          Indicates whether to correct the value if invalid
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bInteger"

   
On Error GoTo ErrHnd
  
   Call M_bValidNumber(vValue, vLow, vHigh, sErrMsg, sCaption, True)
   
   LibValid_bInteger = sErrMsg = ""
   
Exit Function
ErrHnd:
   LibValid_bInteger = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)

End Function
Public Function LibValid_bNetworkPath _
      (ByVal vValue As Variant, ByRef sErrMsg As String, Optional ByVal vMaxLen As Variant) _
       As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments
''' vValue        Path requested to be evaluated
''' sErrMsg       Used to return error information if value is invalid
''' vMaxLen       Optional maximum length of the path
'''
''' sErrMsg       Used to return error information if value is invalid
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bNetworkPath"

   Dim bResult          As Boolean
   Dim sPath            As String
   
On Error GoTo ErrHnd

   sErrMsg = ""
   If LibValid_bPath(vValue, sErrMsg, vMaxLen) Then
      sPath = vValue
      If Left$(sPath, 2) = "\\" Then
         bResult = True
      Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(sPath, gsMUST, gsBE, "a", gsNETWORK, gsPATH)
         Call M_ProcessMsg(sErrMsg, gsNETWORK & " " & gsPATH)
      End If
   End If

   LibValid_bNetworkPath = bResult
   
Exit Function
ErrHnd:
   LibValid_bNetworkPath = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function
Public Function LibValid_bPath _
      (ByVal vValue As Variant, ByRef sErrMsg As String, Optional ByVal vMaxLen As Variant) _
       As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments:
''' vValue        Path requested to be evaluated
''' sErrMsg       Used to return error information if value is invalid
''' vMaxLen       Optional maximum length of the path
'''
'''Description:   Evaluate whether specified is valid directory
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:      True   - If the directory path exists
'''              False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bPath"

   Dim bResult          As Boolean
   Dim sPath            As String
   Dim lAttr            As Long
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   
   '''If specified value is string
   If VarType(vValue) = vbString Then
      sPath = vValue
      '''If path is specified
      If Trim$(sPath) <> "" Then
         
         lAttr = GetAttr(sPath)
         If lAttr And vbDirectory Then
            If Not M_bUseCriteria(vMaxLen) Then
               bResult = True
            Else
               If M_bValidNumber _
                  (vMaxLen, , , sErrMsg, gobjLibStr.LibStr_sBuildSentence(gsMAXIMUM, gsLENGTH)) _
               Then
                  If Len(sPath) <= vMaxLen Then
                     bResult = True
                  Else
                     sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                        (gsLENGTH, gsOF, gsPATH, gsMUST, gsNOT, gsBE, gsMORE, gsTHEN, vMaxLen)
                  End If
               End If
            End If
         Else
            sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPATH, gsMUST, gsBE, "a", gsDIRECTORY)
         End If
      Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPATH, gsIS, gsNOT, gsSPECIFIED)
      End If
   Else
      sErrMsg = gsINVALID & gsDATA & gsTYPE
   End If
   
   Call M_ProcessMsg(sErrMsg, gsPATH)
   
   LibValid_bPath = bResult
   
Exit Function
ErrHnd:
   sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPATH, gsNOT, gsFOUND)
   LibValid_bPath = False
   
Exit Function
End Function

Public Function LibValid_bNumber _
      (ByVal vValue As Variant, _
       Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "" _
      ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments
''' vValue        Value to be evaluated as
''' vLow          minimum value allowed
''' vHigh         maximum value allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bNumber"

On Error GoTo ErrHnd
   
   Call M_bValidNumber(vValue, vLow, vHigh, sErrMsg, sCaption, False)

   LibValid_bNumber = sErrMsg = ""
   
Exit Function
ErrHnd:
   LibValid_bNumber = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sCaption)
   
End Function
Public Function LibValid_bText _
      (ByVal vValue As Variant, _
       Optional ByVal vMinLen As Variant, Optional ByVal vMaxLen As Variant, _
       Optional vValidValues As Variant, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "" _
      ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments
''' vValue        Value to be evaluated as
''' vMinLen       Minimum number of characters
''' vMaxLen       Maximum number of characters
''' vValidValues  Comma delimited list of valid values or an single dimensional array
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified parameters
'''
'''Note:          If list of valid values is not empty and space is not included in the
'''               list of valid values Value will be trimed before length is evaluated
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bText"

   Dim bResult          As Boolean
   Dim lLen             As Long
   Dim sValue           As String
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   
   If sCaption = "" Then
      sCaption = gsVALUE
   End If
   
Err.Clear
On Error GoTo 0
On Error Resume Next

   sValue = vValue
   '''If value is not string compatible
   If Err.Number <> 0 Then
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsTRANSLATE, sCaption, gsINTO, gsTEXT)
   End If
   
On Error GoTo ErrHnd

   If sErrMsg = "" Then
      lLen = Len(sValue)
      If M_bUseCriteria(vMinLen) Then
         If M_bValidNumber _
            (vMinLen, , , sErrMsg, gobjLibStr.LibStr_sBuildSentence(gsMINIMUM, gsLENGTH)) _
         Then
            If lLen < vMinLen Then
               sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                     (gsLENGTH, gsOF, sCaption, gsMUST, gsNOT, gsBE, gsLESS, gsTHEN, vMinLen)
            End If
         End If
      End If
      If sErrMsg = "" Then
         If M_bUseCriteria(vMaxLen) Then
            If M_bValidNumber _
               (vMaxLen, , , sErrMsg, gobjLibStr.LibStr_sBuildSentence(gsMAXIMUM, gsLENGTH)) _
            Then
               If lLen > vMaxLen Then
                  sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                     (gsLENGTH, gsOF, sCaption, gsMUST, gsNOT, gsBE, gsMORE, gsTHEN, vMaxLen)
               End If
            End If
         End If
         If sErrMsg = "" Then
            If M_bUseCriteria(vValidValues) Then
               bResult = LibValid_bListedValue(sValue, vValidValues, , sErrMsg)
            Else
               bResult = True
            End If
         End If
      End If
   End If
   
   Call M_ProcessMsg(sErrMsg, sCaption)

   LibValid_bText = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bText = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)
   
End Function

Public Function LibValid_bTime _
      (ByVal vValue As Variant, _
       Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "" _
      ) As Boolean
'''************************************************************************
'''Called By:     M_bValidVal or Any
'''
'''Arguments:
''' vValue        Value to be evaluated as time
''' vLow          minimum value allowed
''' vHigh         maximum value allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      name of the entity used for error message
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bTime"

On Error GoTo ErrHnd
  
   LibValid_bTime = M_bValidDateTime(vValue, vLow, vHigh, sErrMsg, sCaption, True)
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bTime = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)

End Function
Public Function LibValid_bListedValue _
      (ByVal vValue As Variant, ByVal vValidValues As Variant, Optional ByRef lPosFound As Long, _
       Optional ByVal sCaption As String = "", Optional bEvaluateAsNumber As Boolean, _
       Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     LibValid_bText
'''
'''Arguments
''' vValue        Value been evaluated
''' vValidValues  Comma delimited list of valid values or an single dimensional array
''' lPosFound     Used to return
'''                  starting position of the matching substring if valid values is a string or
'''                  index of the matching element if valid values is an array
''' sCaption      Optional name of the entity used for error message
''' sErrMsg       Used to return error information if value is invalid
''' bEvaluateAsNumber Indicates whether values should be evaluated as number
'''                   e.g. 1 = 1.0 - note: works only if list is an array
'''
'''Description:   Check whether specified value is one of the listed
'''
'''Note:          lPosFound return value of -1 indicates that vValidValues where not specified
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If list of valid values is empty
'''                  True and position found will be set to -1
'''               if found within the list
'''                  True
'''               if Not found or error
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_bListedValue"
   Const lPOS_FOR_EMPTY As Long = -1
   
   Dim bResult          As Boolean
   Dim sValidValues     As String
   Dim lIdx             As Long
   Dim lLen             As Long
   Dim bEmpty           As Boolean
   Dim sValue           As String
   Dim sTestVal         As String
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   
   If VarType(vValidValues) = vbString Then
      sValidValues = vValidValues
      '''If specific values are listed
      If sValidValues <> "" Then
         sValue = vValue
         lLen = Len(sValue)
         '<DEV 08/05/03
         '''Make sure each value is terminated with comma as delimeter
         'sValidValues = gobjLibStr.LibStr_sEndWith(sValidValues, ",", True, True)
         'lPosFound = InStr(sValidValues, sValue)
         '>DEV
         lPosFound = InStr(sValidValues, sValue)
         bResult = lPosFound > 0
      Else
         lPosFound = lPOS_FOR_EMPTY
         bResult = True
      End If
   Else
      ''''''True if found or array is empty
      lPosFound = -1
      bResult = True
      For lIdx = LBound(vValidValues) To UBound(vValidValues)
         bResult = False
         If bEvaluateAsNumber Then
            sValue = vValue
            sTestVal = vValidValues(lIdx)
            bResult = LibValid_dVal(sValue) = LibValid_dVal(sTestVal)
         Else
            bResult = vValue = vValidValues(lIdx)
         End If
         If bResult Then
            lPosFound = lIdx
            Exit For
         End If
      Next lIdx
   End If
   
   If Not bResult Then
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsNOT, gsLISTED, gsVALUE)
   End If
   
   Call M_ProcessMsg(sErrMsg, sCaption)
   

   LibValid_bListedValue = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibValid_bListedValue = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   
End Function
Function LibValid_dRound _
         (ByVal dNumber As Double, ByVal dPrecision As Double, _
          Optional ByVal lUpNearDown As Long = 0 _
         ) As Double
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' dNumber
'''   Number requested to be rounded
''' dPrecision
'''   Rounding Precision. Sign of precision is ignored
''' lUpNearDown
'''   Indicates whether to round Up, Down or to Nearest(Normal Rounding: >= 0.5 - Up, <0.5 - down)
'''   lUpNearDown = 0  Default. Interpreted as request to do normal rounding
'''   lUpNearDown > 0  Interpreted as request to do round up
'''   lUpNearDown < 0  Interpreted as request to do round down
'''
'''Description:   Based on lUpNearDown request round specified number to requested Precision
'''
'''Example:       LibValid_dRound(24.14,1)            = 24
'''               LibValid_dRound(-24.14,1)           = -24
'''               LibValid_dRound(24.574, 0.05)       = 24.55
'''               LibValid_dRound(24.575, 0.05)       = 24.6
'''               LibValid_dRound(-24.575, 0.05)      = -24.6
'''               LibValid_dRound(12.14,0.25,-1)      = 12.00
'''               LibValid_dRound(12.14,0.25,1)       = 12.25
'''               LibValid_dRound(100.12345,0)        = 100.12345
'''               LibValid_dRound(100.12345,0.005,1)  = 100.125
'''
'''Note:          If dPrecision is zero dNumber will not be rounded and returned as is
'''               Rounding is done based on the absolute value
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If Succesful
'''                  dNumber rounded to dPrecision
'''               Else
'''                  Unmodified value
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_dRound"

   Dim dResult    As Double
   Dim dWhole     As Double
   Dim dFraction  As Double
   Dim lSign      As Long
   
On Error GoTo ErrHnd

   If Not dPrecision = 0 Then
      lSign = Sgn(dNumber)
      dNumber = Abs(dNumber)
      dPrecision = Abs(dPrecision)
      
      dWhole = Fix(dNumber / dPrecision)
      dFraction = (dNumber / dPrecision) - dWhole
      dResult = dWhole * dPrecision
      If lUpNearDown = 0 Then          '''Round Normal(nearest)
         '''Compansate - sometimes after division 0.5 is presented as 0.499999999...
         '''So LibValid_dRound(9.35, 0.1) will result in 9.3 and we want 9.4
         If dFraction >= 0.49999999 Then
            dResult = dResult + dPrecision
         End If
      ElseIf lUpNearDown > 0 Then      '''Round Up
         If dFraction > 0 Then
            dResult = dResult + dPrecision
         End If
      '''ElseIf lUpNearDown < 0 Then      '''Round Down
         '''Already determined - do nothing
      End If
      dResult = dResult * lSign  '''Restore Sign
   Else
      dResult = dNumber
   End If
   
   LibValid_dRound = dResult
 
Exit Function
ErrHnd:
   LibValid_dRound = dNumber
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Function LibValid_dValExt(ByVal sStr As String) As Double
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sStr
'''   String expression requested to be converted to and returned as number
'''
'''Description:   If specified string starts with number that can be recognized by CDbl
'''               function convert specified string to number and return it
'''               else return zero
'''
'''Note:          Input/Output differences between This function(LibValid_dVal) and Val() function - its VB couterpart
'''               Numbers expressed in exponential notation:
'''                  Input: "1E1"         / Output: this function 1        - Val function 10
'''                  Input: "9.2D2"       / Output: this function 9.2      - Val function 920
'''               Thousand separator
'''                  Input "1,234.56"     / Output: this function 1234.56  - Val function 1
'''               Dollar Sign
'''                  Input "$1,234.56"    / Output: this function 1234.56  - Val function 0
'''               Debit
'''                  Input "(1,234.56)"   / Output: this function -1234.56 - Val function 0
'''                  Input "$(1,234.56)"  / Output: this function -1234.56 - Val function 0
'''                  Input "($1,234.56)"  / Output: this function -1234.56 - Val function 0
'''
'''               See also LibValid_bIsNum, LibValid_dVal
'''Warning:
'''Updates Globals:
'''
'''Returns:          If specifed string starts with number
'''                     this number
'''                  Else
'''                     Zero
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME                 As String = "LibValid_dValExt"

   Dim dResult                      As Double
   Dim sResult                      As String
   Dim lPos                         As Long
   Dim bOneDot                      As Boolean
   Dim bOpenPrnt                    As Boolean
   Dim bClosePrnt                   As Boolean
   Dim bDollar                      As Boolean
   Dim sChar                        As String
   
   
On Error GoTo ErrHnd

20    sStr = Trim$(sStr)
100   For lPos = 1 To Len(sStr)
120      sChar = Mid$(sStr, lPos, 1)
200      If InStr("0123456789", sChar) > 0 Then
220         sResult = sResult & sChar
300      ElseIf sChar = "." Then
400         If Not bOneDot Then
420            bOneDot = True
440            sResult = sResult & sChar
            Else
500            Exit For
            End If
600      ElseIf lPos > 1 Then
700         If sChar = "," Then
               '''valid if not first char - skip it
800         ElseIf sChar = ")" Then
820            bClosePrnt = True
840            Exit For
1000        ElseIf lPos = 2 Then
1100           If sChar = "$" Then
                  '''if first character was (
1200              If bOpenPrnt Then
1220                 bDollar = True
                  Else
1240                 Exit For
                  End If
1500           ElseIf sChar = "(" Then
                  '''if first character was $
1600              If bDollar Then
1620                 bOpenPrnt = True
                  Else
1640                 Exit For
                  End If
               Else
1800              Exit For
               End If
            Else
1900           Exit For
            End If
2000     Else     '''If lPos = 1 Then
2100        If sChar = "$" Then
               bDollar = True '''valid if starts with $ - skip it
2200        ElseIf sChar = "(" Then
2220           bOpenPrnt = True
            Else
3000            Exit For
            End If
         End If
      Next lPos

9000  If Len(sResult) > 0 Then
9020     dResult = Val(sResult)
9100     If bOpenPrnt Then
9200        If bClosePrnt Then
9220           dResult = -dResult
            Else
               dResult = 0
            End If
         End If
      End If
      
999   LibValid_dValExt = dResult
   
Exit Function
ErrHnd:
   LibValid_dValExt = 0
   Call gobjLibApp.LibApp_Error _
         (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Function LibValid_vCastType _
      (ByVal vValue As Variant, _
       Optional ByVal lPreferredNumeric As VbVarType = vbDouble, _
       Optional ByVal lPreferredUndefined As VbVarType = vbString, _
       Optional ByRef bErr As Boolean _
      ) As Variant
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' vValue
'''   Value requested to be processed
''' lPreferredNumeric
'''   Optional prefered type for numeric values (e.g. vbInteger, vbSingle, vbBoolean, vbString...)
'''   Ignored if value does not have numeric representation. Defaults to Double.
'''   Will cause an error If not set to one of the numeric types
''' lPreferredUndefined
'''   Optional request to return undefined value(Null or Empty) as initial value of specific type
'''   (e.g. vbInteger, vbString, vbEmpty...). Defaults to string
'''   Warning: vbUserDefinedType,vbDataObject,vbArray, vbError, vbObject, vbVariant are not supported
'''
''' bErr
'''   Indicates an error if function fails
'''
'''Description:   Evaluate Data Type of the specified value regardless of representation
'''               Convert return specified value as a variant of the corresponding type
'''               E.g. If its a string representing a date - return as variant of the date type
'''
'''Note:
'''
'''
'''Warning:       This function does not support Objects, arrays, errors...
'''               - For all non-primitives it will always return empty string
'''
'''
'''Updates Globals:
'''
'''Returns:       If Successful
'''                  Specified value as a variant which variant type matches its data type representation
'''               Else
'''                  Non-modified vValue. Also bErr parameter will return true
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibValid_vCastType"
   
On Error GoTo ErrHnd
  
20    bErr = False

      Select Case True
         Case LibValid_bIsNum(vValue)
            Select Case lPreferredNumeric
               Case vbDouble
100               vValue = CDbl(vValue)
               Case vbInteger
110               vValue = CInt(vValue)
               Case vbLong
120               vValue = CLng(vValue)
               Case vbSingle
130               vValue = CSng(vValue)
               Case vbCurrency
140               vValue = CCur(vValue)
               Case vbBoolean
150               vValue = CBool(vValue)
               Case vbDecimal
160               vValue = CDec(vValue)
               Case vbByte
170               vValue = CByte(vValue)
               Case vbString
180               vValue = CStr(vValue)
               Case vbDate
                  vValue = CDate(vValue)
               Case Else
                  bErr = True
200               Call gobjLibApp.LibApp_Error _
                     (1, gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsTYPE, gsSPECIFIED), _
                      msMODULE_NAME, sPROC_NAME, gsTYPE & ": " & lPreferredNumeric _
                     )
            End Select
         Case IsDate(vValue)
300         vValue = CDate(vValue)
         Case IsNull(vValue), IsEmpty(vValue)
            Select Case lPreferredUndefined
               Case vbEmpty
310               vValue = Empty
               Case vbNull
320               vValue = Null
               Case vbInteger
330               vValue = CInt(0)
               Case vbLong
340               vValue = CLng(0)
               Case vbSingle
350               vValue = CSng(0)
               Case vbDouble
360               vValue = CDbl(0)
               Case vbCurrency
370               vValue = CCur(0)
               Case vbDate
380               vValue = CDate(0)
               Case vbString
390               vValue = ""
               Case vbBoolean
400               vValue = False
               Case vbDecimal
410               vValue = CDec(0)
               Case vbByte
420               vValue = CByte(0)
               Case Else   '''vbUserDefinedType, vbDataObject, vbArray, vbError, vbObject, vbVariant
                  bErr = True
500               Call gobjLibApp.LibApp_Error _
                     (1, gobjLibStr.LibStr_sBuildSentence(gsTYPE, gsIS, gsNOT, gsSUPPORTED), _
                      msMODULE_NAME, sPROC_NAME, gsTYPE & ": " & lPreferredUndefined _
                     )
            End Select

         Case IsError(vValue), IsArray(vValue), IsObject(vValue)
800         vValue = ""
         Case Else
900         vValue = CStr(vValue)
      End Select
   
999   LibValid_vCastType = vValue
   
Exit Function
ErrHnd:
   bErr = True
   LibValid_vCastType = vValue
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)
   
Exit Function

End Function

Private Function M_bUseCriteria(ByVal vCriteria As Variant) As Boolean
   '''Evaluate whetner specified validation criteria should be used for evaluation
'''<DEV On Error Resume Next
   '''Use criteria if it is defined else ignore it
   '''<DEV M_bUseCriteria = (Not IsEmpty(vCriteria)) And (Not IsNull(vCriteria)) And (Not IsMissing(vCriteria))
   M_bUseCriteria = Not gobjLibStr.LibStr_sEmptyIfNull(vCriteria) = ""
   
End Function

Public Function M_bValidDateTime _
  (ByVal vValue As Variant, Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
   Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "", _
   Optional bTime As Boolean _
  ) As Boolean
'''************************************************************************
'''Called By:     LibValid_bDateTime
'''               LibValid_bDateTime
'''
'''Arguments:
''' vValue        Value to be evaluated as date
''' vLow          Lowest Date-time allowed
''' vHigh         Highest Data-time allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
''' bTime         Indicates whether to evaluates specified value as a date or a time of the day
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bValidDateTime"
   
   Dim tValue        As Date
   Dim sFormat          As String
   
On Error GoTo ErrHnd
  
   sErrMsg = ""
   
   sFormat = "h:Nn:Ss AM/PM"
   If Not bTime Then
      sFormat = "mm/dd/yyyy " & sFormat
   End If
   If sCaption = "" Then
      sCaption = gsVALUE
   End If
   
   If M_bValidTimeValue(vValue, sErrMsg, bTime, "") Then
      tValue = CDate(vValue)
      If M_bUseCriteria(vLow) Then
         If M_bValidTimeValue(vLow, sErrMsg, bTime, gobjLibStr.LibStr_sBuildSentence(gsMINIMUM, gsRANGE)) Then
            If tValue < CDate(vLow) Then
               sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                        (gsMUST, gsNOT, gsBE, gsBEFORE, Format$(CDate(vLow), sFormat))
            End If
         End If
      End If
      If M_bUseCriteria(vHigh) And (sErrMsg = "") Then
         If M_bValidTimeValue(vHigh, sErrMsg, bTime, gobjLibStr.LibStr_sBuildSentence(gsMAXIMUM, gsRANGE)) Then
            If tValue > CDate(vHigh) Then
               sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                        (gsMUST, gsNOT, gsBE, gsLATER, gsTHAN, Format$(CDate(vHigh), sFormat))
            End If
         End If
      End If
   End If
   
   If sErrMsg <> "" Then
      sErrMsg = sCaption & " - " & sErrMsg
   End If
   
   Call M_ProcessMsg(sErrMsg, sCaption)
   
   M_bValidDateTime = sErrMsg = ""
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   M_bValidDateTime = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)
  
End Function

Public Function M_bValidNumber _
      (ByRef vValue As Variant, _
       Optional ByVal vLow As Variant, Optional ByVal vHigh As Variant, _
       Optional ByRef sErrMsg As String, Optional ByVal sCaption As String = "", _
       Optional ByVal bInteger As Boolean, Optional ByVal lNoOfDigitsAfterDecimal As Long = -1 _
      ) As Boolean
'''************************************************************************
'''Called By:     LibValid_bInteger
'''               LibValid_bNumber
'''Arguments
''' vValue        Value to be evaluated as integer
''' vLow          minimum value allowed
''' vHigh         maximum value allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      name of the entity used for error message
''' bInteger      Indicates whether number must be an integer
''' lNoOfDigitsAfterDecimal
'''   Maximum Number of digits after decimal. If zero or more - specified number will be rounded to it
'''
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bValidNumber"

   Dim dValue           As Double
   Dim dLow             As Double
   Dim dHigh            As Double

On Error GoTo ErrHnd
  
   sErrMsg = ""
   
   If sCaption = "" Then
      sCaption = gsVALUE
   End If
   
   If LibValid_bIsNum(vValue) And Not IsEmpty(vValue) Then
      dValue = vValue
      '''If requested to be rounded
      If lNoOfDigitsAfterDecimal >= 0 Then
         dValue = Round(dValue, lNoOfDigitsAfterDecimal)
         vValue = dValue
      End If
      '''If number must be an integer
      If bInteger Then
         '''If does not contain fraction
         If dValue <> CLng(dValue) Then
            sErrMsg = gobjLibStr.LibStr_sBuildSentence(sCaption, gsMUST, gsBE, "an", gsINTEGER)
         End If
      End If
      If sErrMsg = "" Then
         '''If lower limit is specified
         If M_bUseCriteria(vLow) Then
            If LibValid_bIsNum(vLow) Then
               dLow = vLow
               If dValue < dLow Then
                  sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                           (sCaption, gsMUST, gsNOT, gsBE, gsLESS, gsTHAN, Str$(dLow))
               End If
            Else
               sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsMINIMUM, gsRANGE, gsMUST, gsBE, "a", gsNUMBER)
            End If
         End If
         '''If upper limit is specified
         If sErrMsg = "" And M_bUseCriteria(vHigh) Then
            If LibValid_bIsNum(vHigh) Then
               dHigh = vHigh
               If dValue > dHigh Then
                  sErrMsg = gobjLibStr.LibStr_sBuildSentence _
                           (sCaption, gsMUST, gsNOT, gsBE, gsMORE, gsTHAN, Str$(dHigh))
               End If
            Else
               sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsMAXIMUM, gsRANGE, gsMUST, gsBE, "a", gsNUMBER)
            End If
         End If
      End If
   Else
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(sCaption, gsMUST, gsBE, "a", gsNUMBER)
   End If
   
   M_bValidNumber = sErrMsg = ""
   
Exit Function
ErrHnd:
   M_bValidNumber = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCaption)

End Function

Private Function M_bValidVal _
   (ByVal lType As gnumDTYPDataType, ByVal vValue As Variant, _
    ByVal vMin As Variant, ByVal vMax As Variant, _
    ByRef sErrMsg As String, ByVal sCaption As String _
   ) As Boolean
'''************************************************************************
'''Called By:     LibValid_bValue
'''
'''Arguments:
''' vValue        Value to be evaluated
''' vMin          For string
'''                  Minimum number of characters
'''               For Date and numbers
'''                  Minimum value allowed
'''               For Path and NetworkPath
'''                  Not used
'''               For USPhone
'''                  Minimum length of the phone extension.
'''                  Zero indicates that Extension is optional.
'''                  Greater then zero interpreted as Extension is required
''' vMax          For string
'''                  Maximum number of characters
'''               For Date and numbers
'''                  Maximum value allowed
'''               For Path and NetworkPath
'''                  Not used
'''               For USPhone
'''                  Maximum length of the phone extension
'''                  Zero - indicates that Extension is not allowed
''' sErrMsg       Used to return error information if value is invalid
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   Evaluate whether specified value is of the valid datatype and within
'''               specified range
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bValidVal"

   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   Select Case lType
      Case lDTYP_Long, lDTYP_Integer
         bResult = LibValid_bInteger(vValue, vMin, vMax, sErrMsg, sCaption)
         
      Case lDTYP_Double
         bResult = LibValid_bNumber(vValue, vMin, vMax, sErrMsg, sCaption)
         
      Case lDTYP_Date, lDTYP_DateTime
         bResult = LibValid_bDateTime(vValue, vMin, vMax, sErrMsg, sCaption)
         
      Case lDTYP_String
         bResult = LibValid_bText(vValue, vMin, vMax, "", sErrMsg, sCaption)
         
      Case lDTYP_Path
         bResult = LibValid_bPath(vValue, sErrMsg, vMax)
         
      Case lDTYP_NetworkPath
         bResult = LibValid_bNetworkPath(vValue, sErrMsg, vMax)
         
      Case lDTYP_Boolean
         bResult = LibValid_bBoolean(vValue, sErrMsg, sCaption)
         
      Case lDTYP_Time
         bResult = LibValid_bTime(vValue, vMin, vMax, sErrMsg, sCaption)
      
      Case lDTYP_USPhone
         bResult = LibValid_bUSPhone(vValue, vMin, vMax, sErrMsg, sCaption)
         
      Case lDTYP_InterPhone
         bResult = LibValid_bInterPhone(vValue, sErrMsg, sCaption)
         
      Case Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsSPECIFIED, gsTYPE, gsIS, gsNOT, gsSUPPORTED)
         
   End Select
   
   M_bValidVal = bResult
   
Exit Function
ErrHnd:
   M_bValidVal = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sCaption)
   
End Function
Private Function M_bValidTimeValue _
 (ByVal vValue As Variant, ByRef sErrMsg As String, ByVal bTime As Boolean, _
  ByVal sCaption As String _
 ) As Boolean
'''************************************************************************
'''Called By:     M_bValidDateTime
'''
'''Arguments:
''' vValue        Value to be evaluated as date time or time
''' sErrMsg       Used to return error information if value is invalid
''' bTime         Indicates whether to evaluates specified value as a date or a time of the day
'''
''' sCaption      Optional name of the entity used for error message
'''
'''Description:   if evaluate as time is requested
'''                  Evaluate whether specified value is valid time
'''               else
'''                  Evaluate whether specified value is valid date
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If specified value is valid
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bValidTimeValue"

   Dim tDateTime        As Date
   
On Error GoTo ErrHnd
  
   sErrMsg = ""
   
   If IsDate(vValue) Then
      tDateTime = CDate(vValue)
      If bTime Then
         If CDbl(tDateTime) >= 1 Then
            sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsCONTAINS, gsDATE, gsINFORMATION)
         End If
      Else
         If CDbl(tDateTime) < 1 Then
            sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsDOES, gsNOT, gsCONTAIN, gsDATE, gsINFORMATION)
         End If
      End If
   Else
      sErrMsg = gobjLibStr.LibStr_sBuildSentence _
               (gsCAN, gsNOT, gsBE, gsEVALUATED, gsAS, IIf(bTime, gsTIME, gsDATE))
   End If
   
   If Not (sErrMsg = "") Then
      sErrMsg = Trim$(gobjLibStr.LibStr_sBuildSentence(sCaption, sErrMsg))
   End If
   
   M_bValidTimeValue = (sErrMsg = "")
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   M_bValidTimeValue = False
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")

End Function
Private Sub M_ProcessMsg(ByVal sMsg As String, ByVal sCaption As String)
'''************************************************************************
'''Called By:
'''
'''Arguments
''' sMsg          Contains error information if value is invalid
''' sCaption      Name of the entity used for error message
'''
'''Description:   If invalid value is indicated
'''                  Based on corresponding switches perform or suppress error report action
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_ProcessMsg"

   Dim bResult          As Boolean
   Dim sPath            As String
   
On Error GoTo ErrHnd

   If Not sMsg = "" Then
      If pbLibValidMsgBeep Then
         Call Beep
      End If
      If pbLibValidMsgPop Then
         Call gobjLibMsg.LibMsg_lOKCancel(lMSGOK_INVALID_ENTRY, , , sCaption, sMsg)
      End If
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sCaption & vbCrLf & sMsg)

End Sub

Public Property Get pobjFmt() As nclsLibFormat
   Set pobjFmt = gobjLibFormat
End Property

Private Sub Class_Terminate()
   Set mobjUSPhone = Nothing
End Sub


