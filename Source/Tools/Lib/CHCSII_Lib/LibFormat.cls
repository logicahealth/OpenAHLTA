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
Attribute VB_Name = "nclsLibFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        nclsLibFormat/LibFormat.cls
'''
'''Description:   Common CHCSII format masks and procedures
'''
'''Note:          Exposed as the property of gclsLibValid
'''
'''Modules Required:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "nclsLibFormat."


Public Function DspDate(ByVal vVal As Variant) As String
   DspDate = Format$(vVal, MskDspDate)
End Function


Public Function DspDateTime(ByVal vVal As Variant, Optional ByVal bIncludeSeconds As Boolean = True) As String
   DspDateTime = Format$(vVal, MskDspDateTime(bIncludeSeconds))
End Function


Public Function DspTime(ByVal vVal As Variant, Optional ByVal bIncludeSeconds As Boolean = True) As String
      DspTime = Format$(vVal, MskDspTime(bIncludeSeconds))
End Function

Public Property Get MskDspDate() As String
   MskDspDate = "DD MMM YYYY"
End Property


Public Property Get MskDspDateTime(Optional ByVal bIncludeSeconds As Boolean = True) As String
   MskDspDateTime = "DD MMM YYYY Hh:Nn" & IIf(bIncludeSeconds, ":Ss", "")
End Property


Public Property Get MskDspTime(Optional ByVal bIncludeSeconds As Boolean = True) As String
   MskDspTime = "Hh:Nn" & IIf(bIncludeSeconds, ":Ss", "")
End Property

Public Property Get MskOraDate() As String
   MskOraDate = "MM/DD/YYYY"
End Property


Public Property Get MskOraDateTime(Optional ByVal bIncludeSeconds As Boolean = True) As String
   MskOraDateTime = "MM/DD/YYYY HH:MI" & IIf(bIncludeSeconds, ":SS", "") & " AM"
End Property


Public Property Get MskOraTime(Optional ByVal bIncludeSeconds As Boolean = True) As String
   MskOraTime = "HH:MI" & IIf(bIncludeSeconds, ":SS", "") & " AM"
End Property


Public Property Get MskVBtoOraDate() As String
   MskVBtoOraDate = "mm/dd/yyyy"
End Property

Public Property Get MskVBtoOraDateTime(Optional ByVal bIncludeSeconds As Boolean = True) As String
   MskVBtoOraDateTime = "mm/dd/yyyy hh:mm" & IIf(bIncludeSeconds, ":ss", "") & " AM/PM"
End Property


Public Property Get MskVBtoOraTime(Optional ByVal bIncludeSeconds As Boolean = True) As String
   MskVBtoOraTime = "hh:mm" & IIf(bIncludeSeconds, ":ss", "") & " AM/PM"
    
End Property


Public Function SqlOraDate(ByVal vVal As Variant) As String
   SqlOraDate = "TO_DATE('" & Format$(CDate(vVal), MskVBtoOraDate) & "', '" & MskOraDate & "')"
End Function

Public Function SqlOraTime(ByVal vVal As Variant) As String
   SqlOraTime = "TO_DATE('" & Format$(CDate(vVal), MskVBtoOraTime) & "', '" & MskOraTime & "')"
End Function

Public Function SQLQuote(ByVal sValue As String) As String
'''************************************************************************
'''Called by:     SQLQuoteList
'''
'''Arguments:
''' sValue        String values requested to be formated for SQL statement
'''
'''Description:
''' doubles up the single quotes and wraps the string in single quotes
''' e.g.:     Men's club of "Greater" NY
''' becomes: 'Men''s club of ""Greater"" NY'
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''Returns:       If successful
'''                  formated values
'''               Else
'''                  '' - representation of an empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "SQLQuote"
   
   Dim sResult       As String
   
On Error GoTo ErrHnd

100   If Len(sValue) > 0 Then
120      sResult = "'" & Replace(Replace(sValue, Chr(34), Chr(34) & Chr(34)), "'", "''") & "'"
      Else
500      sResult = "''"
      End If
   
999   SQLQuote = sResult
   
Exit Function
ErrHnd:
   SQLQuote = "''"
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function SQLQuoteList(ParamArray sValues()) As String
'''************************************************************************
'''Arguments:
''' sValues       Array of string values
'''
'''Description:   Convert specified list to comma-delimited
'''               list of string values formated for SQL insert statement
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''Returns:       If successful
'''                  formated list
'''               Else
'''                  empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "SQLQuoteList"
   
   Dim sResult          As String
   Dim lIdx             As Long
   
On Error GoTo ErrHnd
   
100   For lIdx = LBound(sValues) To UBound(sValues)
140      sResult = sResult & SQLQuote(sValues(lIdx)) & ","
      Next lIdx
   
      '''Remove trailing comma
999   SQLQuoteList = gobjLibStr.LibStr_sEndWith(sResult, ",", False, False)
   
Exit Function
ErrHnd:
   SQLQuoteList = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function VBtoOraDate(ByVal vVal As Variant) As String
   VBtoOraDate = Format$(vVal, MskVBtoOraDate)
End Function

Public Function VBtoOraDateTime(ByVal vVal As Variant) As String
   VBtoOraDateTime = Format$(vVal, MskVBtoOraDateTime)
End Function


Public Function VBtoOraTime(ByVal vVal As Variant) As String
   VBtoOraTime = Format$(vVal, MskVBtoOraTime)
End Function


Public Function SqlOraDateTime(ByVal vVal As Variant) As String
   SqlOraDateTime = "TO_DATE('" & Format$(CDate(vVal), MskVBtoOraDateTime) & "', '" & MskOraDateTime & "')"
End Function


