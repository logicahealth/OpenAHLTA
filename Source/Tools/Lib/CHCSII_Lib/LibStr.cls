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
Attribute VB_Name = "gclsLibStr"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibStr/LibStr.cls
'''
'''Description:   String processing utilities and associated entities.
'''
'''Modules Required:
'''   basLibApp
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "gclsLibStr."

Public Enum pnumTXFTextFormat
   TxfUndefined    '''0 - will cause an error   '<ENHANCE use to autodetect source format
   TxfWinANSI      '''1 - Text in Windows ANSI format (an end of a paragraph is marked with the control characters 13 and 10).
   TxfANSI         '''2 - Text in ANSI format (an end of a paragraph is marked only with the control character 10).
   TxfTX           '''3 - Text including formatting attributes in the internal Text Control format. Text is stored in ANSI.
   TxfHTML         '''4 - HTML format (Hypertext Markup Language).
   TxfRTF          '''5 - RTF format (Rich Text Format).
   TxfUnicodeWin   '''6 - Text in Windows Unicode format (an end of a paragraph is marked with the control characters 13 and 10).
   TxfUnicode      '''7 - Text in Unicode format (an end of a paragraph is marked only with the control character 10).
   TxfUnicodeTX    '''8 - Text including formatting attributes in the internal Text Control format. Text is stored in Unicode.
   TxfWord         '''9 - Microsoft Word format.
   TxfXML          ''10 - XML format (Extensible Markup Language).
   TxfCSS          ''11 - CSS format (Cascading Style Sheet).
End Enum

Public Function LibStr_bLettersOnly _
      (ByVal sSource As String, ByRef lPosInvalid As Long, Optional ByVal bSpacesAllowed = True) _
       As Boolean
'''***********************************************************************
'''Arguments:
''' sSource
'''   String requested to be evaluated
''' lPosInvalid
'''   Position of the non-letter character within source string. Zero if none found
''' bSpacesAllowed
'''   Indicates whether space should be evaluated as valid character
'''
'''Description:   Evaluate whether specified string consists of
'''               letter and space characters only.
'''
'''Uses Modules:  None
'''
'''Uses Globals:  None
'''
'''Returns:       If Test string contains only valid characters  or source is empty
'''                  True
'''               Else or error
'''                  False
'''***********************************************************************
   Const sPROC_NAME   As String = msMODULE_NAME & "LibStr_bLettersOnly"
   
   Dim bResult       As Boolean
   Dim lIdx          As Long
   Dim lAsc          As Long
   
On Error GoTo ErrHnd

   lPosInvalid = 0
   bResult = True
   sSource = UCase$(sSource)
   
   '''Keep evaluating until ens of the string or first invalid character is found
   For lIdx = 1 To Len(sSource)
      lAsc = Asc(Mid$(sSource, lIdx, 1))
      If lAsc < 65 Then
         If lAsc = 32 Then
            bResult = bSpacesAllowed
         Else
            bResult = False
         End If
      ElseIf lAsc > 90 Then
         bResult = False
      End If
      If Not bResult Then
         lPosInvalid = lIdx
         Exit For
      End If
   Next lIdx
   
   LibStr_bLettersOnly = bResult
   
Exit Function
ErrHnd:
   LibStr_bLettersOnly = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function


Public Function LibStr_bValidString _
        (ByVal sTestStr As String, ByVal sEvalChars As String, _
         Optional ByVal bEvalCharsAsValid As Boolean = False, _
         Optional bIgnoreCase As Boolean, Optional ByRef lPosInvalid As Long _
        ) As Boolean
'''***********************************************************************
'''Called by:     Any
'''
'''Arguments:
''' sTestStr    - string to evaluate
''' sEvalChars - String containg characters that based on bEvalCharsAsValid are considered to be
'''                         valid or invalid part of the test string
''' bEvalCharsAsValid - Indicates whether sEvalChars contains Valid or invalid character set
'''
''' lPosInvalid
'''   Position of the non-letter character within source string. Zero if none found
'''
'''Description:   This procedure tests whether or not the passed Test string
'''               contains valid characters only.
'''
'''Uses Modules:  None
'''
'''Uses Globals:  None
'''
'''Returns:       True if Test string contains only valid characters
'''               False - otherwise
'''***********************************************************************
   Const sPROC_NAME   As String = msMODULE_NAME & "LibStr_bValidString"
   
   Dim bResult       As Boolean
   Dim lCharNo       As Long
   Dim bFound        As Boolean
   
On Error GoTo ErrHnd

   bResult = True '''Will be reset if invalid
   
   '''If there is anything to evaluate
   If Not (sTestStr = "") Then
      If bIgnoreCase Then
         sTestStr = UCase$(sTestStr)
         sEvalChars = UCase$(sEvalChars)
      End If
      
      '''Do until all characters are eveluated or one fails
      Do While (lCharNo < Len(sTestStr)) And bResult
         lCharNo = lCharNo + 1
         bFound = InStr(sEvalChars, Mid$(sTestStr, lCharNo, 1)) > 0
         
         If bEvalCharsAsValid Then
            bResult = bFound
         Else
            bResult = Not bFound
         End If
      Loop
   Else
      If bEvalCharsAsValid Then
         bResult = sEvalChars = ""
      Else
         bResult = False
      End If
   End If
   
   If Not bResult Then
      lPosInvalid = lCharNo
   End If
   
   LibStr_bValidString = bResult
   
Exit Function
ErrHnd:
   LibStr_bValidString = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function
Public Function LibStr_lInStrNotPrecededBy _
      (ByVal sSource As String, ByVal sMatch As String, ByVal sCharsPrecededBy As String, _
       Optional ByRef sPrecededChar As String, Optional ByRef lPrecededCharPos As Long, _
       Optional ByVal lCompareMtd As VbCompareMethod = vbTextCompare, _
       Optional ByVal lStartPosition As Long = 1 _
      ) As Long
'''************************************************************************
'''Arguments:
''' sSource
'''   String being searched
''' sMatch
'''   String sought
''' sCharsPrecededBy
'''   List of characters to seek before sMatch
''' sPrecededChar
'''   Preceded character. Empty if not found. Returned byref.
''' lPrecededCharPos
'''   Position of the first occurrence of the preceded character within source string.
'''   Zero if not found. Returned byref.
''' lCompareMtd
'''   Indicates type of string comparison to perform. Defaults to vbTextCompare - not case sensitive
''' lStartPosition
'''   Optional starting position of the search, Default to 1
'''
'''Description:   Search specified source string for sMatch.
'''               If match is found is is confirmed only if none of the characters
'''               specified in sCharsPrecededBy are found before position of the match
'''
'''Note:          Case sensitive
'''
'''Examples:
'''   LibStr_lInStrNotPrecededBy("""A, B", ",", """'", sChr, lPos)
'''      Will Return 0, sChr ", lPos 1
'''   LibStr_lInStrNotPrecededBy("Ax B", "x", """'", sChr, lPos)
'''      Will Return 2, sChr <Empty>, lPos 0
'''   LibStr_lInStrNotPrecededBy("X11A11B", "B", "1X", sChr, lPos)
'''      Will Return 0, sChr X, lPos 1
'''   LibStr_lInStrNotPrecededBy("1X1A11B", "B", "X1", sChr, lPos)
'''      Will Return 0, sChr 1, lPos 1
'''   LibStr_lInStrNotPrecededBy("123xM6M8x01M3", "M", "x", sChr, lPos, , 8)
'''      Will Return 0, sChr x, lPos 9
'''   LibStr_lInStrNotPrecededBy("123xM6M8x01M3", "M", "x", sChr, lPos, , 10)
'''      Will Return 12, sChr <Empty>, lPos 0
'''
'''Note:
'''
'''Warning:
'''
'''Returns:       If sSource or sMatch is Empty
'''                  Zero
'''               if Match found and none of the characters
'''               specified in sCharsPrecededBy are found before position of the match
'''                  Position of the match
'''               else
'''                  Zero
'''               if error
'''                  -1
'''
'''Author:        Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_lInStrNotPrecededBy"

   Dim lResult             As Long
   Dim lMatchPos           As Long
   Dim lIdx                As Long
   
On Error GoTo ErrHnd

   sPrecededChar = ""
   lPrecededCharPos = 0
   If Len(sSource) > 0 And Len(sMatch) > 0 Then
10    lMatchPos = InStr(lStartPosition, sSource, sMatch, lCompareMtd)
      If lMatchPos > 0 Then
         For lIdx = lStartPosition To lMatchPos
            '''Search for preceded characters and If one found
20          If InStr(1, sCharsPrecededBy, Mid$(sSource, lIdx, 1), lCompareMtd) > 0 Then
               '''Report it to the caller
               lPrecededCharPos = lIdx
30             sPrecededChar = Mid$(sSource, lIdx, 1)
               Exit For
            End If
         Next lIdx
         '''If none of the preceded characters are found
         If lPrecededCharPos = 0 Then
            lResult = lMatchPos
         '''otherwise disqualify the match
         End If
      End If
   End If
   
   LibStr_lInStrNotPrecededBy = lResult
   
Exit Function
ErrHnd:
   LibStr_lInStrNotPrecededBy = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function LibStr_lInStrPlus _
(ByVal sSource, ByVal sCharBefore As String, ByVal sMatch As String, ByVal sCharAfter As String, _
 Optional ByVal lCompareMtd As VbCompareMethod = vbTextCompare, Optional ByVal lStartPosition As Long = 1 _
) As Long
'''************************************************************************
'''Arguments:
''' sSource
'''   String being searched
''' sCharBefore
'''   List of characters that can occur immediately before to qualify the match
''' sMatch
'''   String sought
''' sCharAfter
'''   List of characters that can occur immediately after to qualify the match
''' lCompareMtd
'''   Indicates type of string comparison to perform. Defaults to vbTextCompare - not case sensitive
''' lStartPosition
'''   Optional starting position of the search, Default to 1
'''
'''Description:   Search specified source string for string matching sMatch.
'''               If character immediately presiding the match is one of the character
'''               specified in sCharBefore OR sMatch is at the beginning of the Source
'''               AND character immediately following sMatch is one of the characters
'''               specified in sCharAfter OR sMatch is at the end of the Source
'''               Return position of the Match else search for next occurrence.
'''
'''Example:       LibStr_lInStrPlus ("123.ABC,456", ".", "ABC", ",")    will return 5
'''               LibStr_lInStrPlus ("123.ABC.456", ".,x", "ABC", "x")  will return 0
'''               LibStr_lInStrPlus ("123.ABC.456", ".,", "ABC", "")    will return 5
'''               LibStr_lInStrPlus ("ABC.123", ".,", "ABC", ".,")      will return 1
'''               LibStr_lInStrPlus ("123.ABC", ".,", "ABC", ".,")      will return 5
'''               LibStr_lInStrPlus ("123.ABC", "", "ABC", "")          will return 5
'''
'''Note:
'''
'''Warning:
'''
'''Returns:       If sSource or sMatch is Empty
'''                  Zero
'''               if Match found
'''                  If match is in the midle of the source
'''                     if sCharBefore and sCharAfter are not empty
'''                        If character immediately presiding the match is one of the
'''                        characters specified in sCharBefore
'''                        AND character immediately following sMatch is one of the
'''                        characters specified in sCharAfter
'''                           *Starting POSITION of the match
'''                     if sCharBefore is empty or match is at the beginning of the source
'''                     and sCharAfter is not empty
'''                        If character immediately following sMatch is one of the
'''                        characters specified in sCharAfter
'''                           *Starting POSITION of the match
'''                     if sCharBefore is not empty
'''                     and sCharAfter is empty or or match is at the end of the source
'''                        If character immediately presiding the match is one of the
'''                        characters specified in sCharBefore
'''                           *Starting POSITION of the match
'''                     if sCharBefore and sCharAfter are empty
'''                        *Starting POSITION of the match
'''               if Match not found
'''                  Zero
'''               If error
'''                  -1
'''
'''Author:        Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_lInStrPlus"

   Dim lResult             As Long
   Dim bFound              As Boolean
   Dim lPos                As Long
   Dim lLenSource          As Long
   Dim lLenMatch           As Long
   Dim bEvalBefore         As Boolean
   Dim bEvalAfter          As Boolean
   
On Error GoTo ErrHnd

   lLenSource = Len(sSource)
   lLenMatch = Len(sMatch)
   If lLenSource > 0 And lLenMatch > 0 Then
20    lPos = InStr(lStartPosition, sSource, sMatch, lCompareMtd)
   End If
   
   Do While (lPos > 0) And (Not bFound)
      bEvalBefore = False
      bEvalAfter = False
      '''If match is in the midle of the source
      If (lPos > 1) And ((lPos + lLenMatch - 1) < lLenSource) Then
100      bEvalBefore = Len(sCharBefore) > 0  '''Evaluate char before if requested
         bEvalAfter = Len(sCharAfter) > 0    '''Evaluate char after if requested
      '''If Match is at the begining of the source
      ElseIf (lPos = 1) Then
         '''If there is anything after the match
200      If ((lPos + lLenMatch - 1) < lLenSource) Then
220         bEvalAfter = Len(sCharAfter) > 0    '''Evalaute character after only - if requested
         '''Match is equal to source
         '''(lPos = 1) and (lPos + lLenMatch - 1) = lLenSource
         End If
      Else
         '''Match is at the end of the source line and not equal to source
         '''(lPos > 1) and lPos + lLenMatch - 1) = lLenSource
300      bEvalBefore = Len(sCharBefore) > 0  '''Evaluate character before only - if requested
      End If
      '''If character before requested to be evaluated
      If bEvalBefore Then
         '''If character before qualifies the match
500      If InStr(1, sCharBefore, Mid$(sSource, lPos - 1, 1), lCompareMtd) > 0 Then
            bFound = True
         Else
            bEvalAfter = False   '''Not need to check - already failed
         End If
      ElseIf Not bEvalAfter Then
         bFound = True
      End If
      '''If character after need to be evaluated
      If bEvalAfter Then
         '''And character after qualifies the match
600      If InStr(1, sCharAfter, Mid$(sSource, lPos + lLenMatch, 1), lCompareMtd) > 0 Then
            bFound = True
         Else
            bFound = False '''Could be set to true in eval before
         End If
      End If
      '''If match was found
      If bFound Then
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Success
         lResult = lPos
         Exit Do
      Else
         '''Look for next match
         lPos = lPos + 1
800      lPos = InStr(lPos, sSource, sMatch, lCompareMtd)
      End If
   Loop
   
   LibStr_lInStrPlus = lResult
   
Exit Function
ErrHnd:
   LibStr_lInStrPlus = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function LibStr_lInStrWithin _
(ByVal sSource As String, ByVal sMatch As String, ByVal sWithinChar As String, ByVal bWithin, _
 Optional ByVal lCompareMtd As VbCompareMethod = vbTextCompare, _
 Optional ByVal lStartPosition As Long = 1 _
) As Long
'''************************************************************************
'''Arguments:
''' sSource
'''   String being searched
''' sMatch
'''   String sought
''' sWithinChar
'''   Specified Delimiter
''' bWithin
'''   Indicates whether to search outside or within of the specified delimiters
''' lCompareMtd
'''   Indicates type of string comparison to perform. Defaults to vbTextCompare - not case sensitive
''' lStartPosition
'''   Optional starting position of the search, Default to 1
'''
'''Description:   Search specified source string for sMatch.
'''               If sMatch found and not Preceded and followed by sWithin
'''                  Return its position
'''               else
'''                  Search for next occurrence.
'''
'''Example:       LibStr_lInStrWithin ("|12|.|2|.|3|", "2", "|", True)           will return 3
'''               LibStr_lInStrWithin ("|12|.|2|.|3|", "2", "|", False)          will return 0
'''               LibStr_lInStrWithin ("|ABC|..ABC..|3|", "ABC", "|", False)     will return 8
'''               LibStr_lInStrWithin ("|ABC|..ABC..|3|", "ABC", "|", True)      will return 2
'''
'''Note:
'''               LibStr_lInStrWithin ("|ABC|..ABC..|3|", "ABC", "|", False,,3)  will return 0
'''               because starting from position 3 "ABC" is within "|" >> "ABC|..ABC..|"
'''               However
'''               LibStr_lInStrWithin ("|ABC|..ABC..|3|", "ABC", "|", False,,2)  will return 2
'''               because starting from position 2 "ABC" is not within "|" >> "ABC|"
'''Warning:
'''               Only first (most left) character of the sWithinChar argument will be used as delimiter
'''                  the rest of the string will be ignored
'''               if sMatch contains character(s) specified in sWithinChar
'''                  function will return 0
'''
'''Returns:       If sSource or sMatch is Empty
'''                  Zero
'''               if Match found
'''                  If within requested (bWithin is true)
'''                     if match is within characters specified in sWithinChar argument
'''                        Position of the match
'''                     else
'''                        Zero
'''                  If not within requested (bWithin is false)
'''                     if match is not within characters specified in sWithinChar argument
'''                        Position of the match
'''                     else
'''                        Zero
'''               else
'''                  Zero
'''               If error
'''                  -1
'''
'''Author:        Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_lInStrWithin"

   Dim lResult             As Long
   Dim lPos                As Long
   Dim lLenSource          As Long
   Dim lLenMatch           As Long
   Dim bWithinON           As Boolean
   
On Error GoTo ErrHnd

2     lLenSource = Len(sSource)
4     lLenMatch = Len(sMatch)
6     lPos = lLenSource '''Default to Cancel
      '''If initial string contain data
      If lLenSource > 0 And lLenMatch > 0 Then
10       sWithinChar = Left$(sWithinChar, 1)    '''Use only first character of the specified delimiter
20       If lCompareMtd = vbTextCompare Then
22          sWithinChar = UCase$(sWithinChar)
24          sMatch = UCase$(sMatch)
26          sSource = UCase$(sSource)
         End If
         '''sMatch does not contain character(s) specified in sWithinChar
40       If InStr(sMatch, sWithinChar) = 0 Then
            '''Initiate starting position
            lPos = lStartPosition - 1
         '''Else Cancel and return Zero
         End If
      '''Else Cancel and return Zero
      End If
      '''If there is anything to look for and valid initial values - scan the source string
      '''Search until match is found or all characters of the source evaluated
      Do While (lPos < lLenSource) And (lResult = 0)
         lPos = lPos + 1
100      Select Case Mid$(sSource, lPos, 1)
            Case sWithinChar
200            bWithinON = Not bWithinON  '''Process delimiter in pairs - Toggle within state
            Case Else
               '''If rest of the source has enough room to fit the match
300            If (lPos + lLenMatch - 1) <= lLenSource Then
                  '''If within requested and within or not within requested and not within
400               If (bWithin And bWithinON) Or ((Not bWithin) And (Not bWithinON)) Then
                     '''If match found
500                  If Mid$(sSource, lPos, lLenMatch) = sMatch Then
600                     lResult = lPos '''Return position of the match and exit the loop
                     End If
                  End If
               End If
         End Select
      Loop

999   LibStr_lInStrWithin = lResult
   
Exit Function
ErrHnd:
   LibStr_lInStrWithin = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function LibStr_lOccurrencePos _
      (ByVal sSource As String, ByVal sDelimiter As String, ByVal lOccurrence As Long, _
       Optional ByVal bCountLeftToRight As Boolean = True _
      ) As Long
'''************************************************************************
'''Arguments:
''' sSource
'''   String being searched
''' sDelimiter
'''   Specified delimiter
''' lOccurrence
'''   Indicates occurance number of the specified delimiter
''' bCountLeftToRigth
'''   Indicates whether to count from the begining(left to right) or end(right to left)
'''   of the source string. Defaults to left to right
'''
'''Description:   Based on requested direction of the count(bCountLeftToRigth switch)
'''               find starting position of requested occurance of the specified delimiter
'''               within specified source string
'''
'''Note:          Case sensitive
'''
'''Example:       LibStr_lOccurrencePos ("1.3.5",".",2)          will return 4
'''               LibStr_lOccurrencePos ("1.3.5",".",2,false)    will return 2
'''               LibStr_lOccurrencePos ("123",".")              will return 0
'''Note:
'''Warning:       If delimiter is empty function will be return 1
'''
'''Returns:       If Delimeter found
'''                  Starting Position
'''               if Source string is empty
'''               or Delimiter is not found
'''               or requested Occurrence < 1
'''                  0
'''               if Delimiter is empty
'''                  1
'''               if Delimiter occur fewer times then specified
'''                  -1
'''               if Error
'''                  -2
'''
'''Author:        Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_lOccurrencePos"

   Dim lResult             As Long
   Dim lPos                As Long
   Dim lLenDel             As Long
   Dim lOccurrenceCount    As Long
   
On Error GoTo ErrHnd

   lLenDel = Len(sDelimiter)
   
   If Len(sSource) > 0 Then
      If lLenDel > 0 Then
         
         If bCountLeftToRight Then
            lPos = 1
         Else
            lPos = -1
         End If
         
         Do
            If bCountLeftToRight Then
               lPos = InStr(lPos, sSource, sDelimiter)
            Else
               lPos = InStrRev(sSource, sDelimiter, lPos)
            End If
            '''If found
            If lPos > 0 Then
               lOccurrenceCount = lOccurrenceCount + 1
               '''If requested occurance
               If lOccurrenceCount = lOccurrence Then
                  '''Store it and exit the loop '''''''''''''''''''''''''''''''''''''''''
                  lResult = lPos
               Else
                  '''Try to find next occurance
                  If bCountLeftToRight Then
                     lPos = lPos + lLenDel
                  Else
                     lPos = lPos - lLenDel
                  End If
               End If
            End If
         Loop Until (lResult > 0) Or (lPos < 1)
         '''Search until found or end of the string
         
         '''If Delimiter occur fewer times then specified
         If lResult = 0 And lOccurrenceCount > 0 Then
            '''Indicate this condition '''''''''''''''''''''''''''''''''''''''''''''''''
            lResult = -1
         End If
      Else
         '''Delimiter is empty ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         lResult = 1
      End If
   End If
   
   LibStr_lOccurrencePos = lResult
   
Exit Function
ErrHnd:
   LibStr_lOccurrencePos = -2
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function
Public Function LibStr_sBeforeFirstDelimiter _
      (ByVal sSource As String, ByVal sDelimiter As String, _
       Optional ByRef lPosAfter As Long, Optional ByRef bErr As Boolean _
      ) As String
'''************************************************************************
'''Arguments:
''' sSource
'''   String being searched
''' sDelimiter
'''   Specified delimiter
''' lPosAfter
'''   Returns first position within sSource after first occurance of the sDelimiter
''' bErr
'''   Indicates whether err occcured
'''
'''Description:   Return part of the string before the first occurance of the
'''               specified delimeter
'''
'''Note:          Case sensitive
'''               See also LibStr_sAfterLastDelimiter
'''
'''Examples:      LibStr_sBeforeFirstDelimiter ("A.B.C",".")       will return A
'''               LibStr_sBeforeFirstDelimiter ("A.B..C","..")     will return A.B
'''               LibStr_sBeforeFirstDelimiter ("A.B.C",",")       will return A.B.C
'''
'''
'''Warning:       If delimiter is empty or not found function will be return sSource
'''               if delimeter is at the end of the string
'''                  lPosAfter will be set to len of the source plus one
'''               If delimiter is at the begining of the source
'''                  Function will return empty string
'''
'''Returns:          If delimeter is found
'''                     String before first occurance of the delimeter
'''                  Else
'''                     Source
'''                  If Error
'''                     Empty string and bErr will be set to true
'''
'''Author:        Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_sBeforeFirstDelimiter"

   Dim sResult             As String
   Dim lPos                As Long
   
On Error GoTo ErrHnd

   bErr = False
   lPosAfter = 0
   sResult = sSource
   
   If Len(sDelimiter) > 0 Then
      lPos = InStr(1, sSource, sDelimiter)
      If lPos > 0 Then
         sResult = Left$(sSource, lPos - 1)
         lPosAfter = lPos + Len(sDelimiter)
      End If
   End If
   
   LibStr_sBeforeFirstDelimiter = sResult
   
Exit Function
ErrHnd:
   LibStr_sBeforeFirstDelimiter = ""
   bErr = True
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function LibStr_sBuildDelimited(ByVal sDelimiter As String, ParamArray vElements()) As String
'''************************************************************************
'''Arguments:
''' sDelimiter    specified delimiter
''' vElements()   Array of elements to be combined into delimited string
'''
'''Description:   Convert specified array of elements into string delimited by specified
'''               delimiter
'''Note:
'''
'''Warning:       Elements must contain string compatible data types
'''
'''Updates Globals:
'''Returns:       If successful
'''                  Delimited string
'''               Else
'''                  empty
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_sBuildDelimited"
   
   Dim sResult       As String
   Dim lIdx          As Long
   
On Error GoTo ErrHnd
   
   For lIdx = LBound(vElements) To UBound(vElements)
      sResult = sResult & sDelimiter & vElements(lIdx)
   Next lIdx
   
   '''Return result
   LibStr_sBuildDelimited = Mid$(sResult, Len(sDelimiter) + 1)
   
Exit Function
ErrHnd:
   LibStr_sBuildDelimited = sResult
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsINDEX & lIdx)
   
End Function
Public Function LibStr_sBuildTitle(ParamArray sWords()) As String
'''************************************************************************
'''Arguments:
''' sWords        Array of words to be combined into Sentence
'''
'''Description:   Process specified group of words
'''               Convert all words to lower Case
'''               Capitalize all Words and
'''               Remove accessed spaces
'''Note:
'''
'''Warning:       sWords must contain at least one word with at least one
'''               not space character
'''
'''Updates Globals:
'''Returns:       If successful
'''                  Sentence
'''               Else
'''                  Empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_sBuildTitle"
   
   Dim sResult       As String
   Dim lIdx          As Long
   Dim sWord         As String
   
On Error GoTo ErrHnd
   
   For lIdx = LBound(sWords) To UBound(sWords)
      sWord = Trim$(sWords(lIdx))
      '''If all caps - use it as is
      If Not UCase$(sWord) = sWord Then
         sWord = LCase$(sWords(lIdx))
         '''Always capitalize ID
         If sWord = "id" Then
            sWord = UCase$(sWord)
         '''If first word
         ElseIf lIdx = 0 Then
            '''Capitalize first letter
            Mid$(sWord, 1, 1) = UCase$(Left$(sWord, 1))
         ElseIf Not InStr("to,from,for,in,and,or,of,as", sWord) > 0 Then
            '''Capitalize first letter
            Mid$(sWord, 1, 1) = UCase$(Left$(sWord, 1))
         End If
      End If
      '''Omit space between word and (s)
      If sWord = "(s)" Then
         sResult = sResult & sWord
      Else
         sResult = sResult & " " & sWord
      End If
   Next lIdx
   
   '''Return result
   LibStr_sBuildTitle = Trim$(sResult)
   
Exit Function
ErrHnd:
   LibStr_sBuildTitle = sResult
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function LibStr_sAfterLastDelimiter _
               (ByVal sSource As String, ByVal sDelimiter As String, _
                Optional ByRef lPosAfter As Long, Optional ByRef lNoOfOccurrences As Long, _
                Optional ByRef bErr As Boolean _
               ) As String
'''************************************************************************
'''Arguments:
''' sSource
'''   String being searched
''' sDelimiter
'''   Specified delimiter
''' lPosAfter
'''   Returns first position within sSource after last occurance of the sDelimiter
''' lNoOfOccurrences
'''   Returns number of occurrences of the specified sDelimiter
''' bErr
'''   Indicates whether err occcured
'''
'''Description:   Return part of the string after the last occurance of the
'''               specified delimeter
'''
'''Note:          Case sensitive
'''
'''Examples:      LibStr_sAfterLastDelimiter ("A.B.C",".")       will return C
'''               LibStr_sAfterLastDelimiter ("A..B.C","..")     will return B.C
'''               LibStr_sAfterLastDelimiter ("A.B.C",",")       will return A.B.C
'''
'''Note:          See also LibStr_sBeforeFirstDelimiter
'''
'''Warning:       If delimiter is empty or not found function will be return sSource
'''               if delimeter is at the end of the string
'''                  lPosAfter will be set to len of the source plus one
'''                  And function will return empty string
'''
'''Returns:          If delimeter found
'''                     String after last occurance of the delimeter
'''                  Else
'''                     Source
'''                  If Error
'''                     Empty string and bErr will be set to true
'''
'''Author:        Alexander Alberstein
'''
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_sAfterLastDelimiter"

   Dim sResult             As String
   Dim lPos                As Long
   Dim lPosLast            As Long
   Dim lLenDel             As Long
   
On Error GoTo ErrHnd

   bErr = False
   lPosAfter = 0
   lNoOfOccurrences = 0
   sResult = sSource
   lLenDel = Len(sDelimiter)
   
   If lLenDel > 0 Then
      '''Find last occurance of the specified delimeter
      '''Note: InStrRev is not used because we want to count number of occurrences
      lPos = 1
      Do
         lPos = InStr(lPos, sSource, sDelimiter)
         If lPos > 0 Then
            '''Try ti find next occurance
            lPos = lPos + lLenDel
            lPosLast = lPos
            lNoOfOccurrences = lNoOfOccurrences + 1
         End If
      Loop Until lPos = 0
      '''If delimeter was found
      If lPosLast > 0 Then
         sResult = Mid$(sSource, lPosLast)
         lPosAfter = lPosLast
      End If
   End If
   
   LibStr_sAfterLastDelimiter = sResult
   
Exit Function
ErrHnd:
   LibStr_sAfterLastDelimiter = ""
   bErr = True
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function LibStr_sConvertText _
      (ByVal vSource As Variant, _
       ByVal lSourceFormat As pnumTXFTextFormat, ByVal lReturnFormat As pnumTXFTextFormat _
      ) As String
'''************************************************************************
'''Arguments:
''' vSource
'''   Text requested to be converted.  Must be a string or a byte array
''' lSourceFormat
'''   Format of the specified text
''' lReturnFormat
'''   Format to which specified text was requested to be converted
'''
'''Description:   Convert specifed string to requested format
'''
'''
'''
'''Note:
'''
'''Warning:
'''
'''
'''Updates Globals:
'''Returns:       If successful
'''                  Converted text
'''               Else
'''                  Empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_sConvertText"
   
   Dim sResult          As String
   Dim frmLibCnt        As frmLibCnt
   Dim yBuffer()        As Byte
   
On Error GoTo ErrHnd
   
100   Set frmLibCnt = P_frmLibCnt
      
200   With frmLibCnt.txoLib
         '''Clear control
220      Call .LoadFromMemory(" ", TxfANSI)
         '''If there is anything to convert - avoid errors on empty HTML, RTF ...
300      If Len(vSource) > 0 Then
            '''Load source string
320         Call .LoadFromMemory(vSource, lSourceFormat)
         End If
         '''Return source in requested format
         '''Word format causes an error if a string is used as a buffer type identifier
500      If Not lReturnFormat = TxfWord Then
520         sResult = .SaveToMemoryBuffer(" ", lReturnFormat)
         Else
540         sResult = .SaveToMemoryBuffer(yBuffer, lReturnFormat)
         End If
      End With
      
700   If lReturnFormat = TxfANSI Then
720      sResult = LibStr_sReplaceChar(sResult, Chr$(11), 0, False, " ")
      End If
      
999   LibStr_sConvertText = Trim$(sResult)
   
Exit Function
ErrHnd:
   LibStr_sConvertText = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function LibStr_sExtractSubstring _
      (ByVal sLine As String, ByRef sBefore As String, ByRef sAfter As String, Optional ByVal sMidle As String _
      ) As String
'''************************************************************************
'''Arguments:
''' sLine         String containing all three of the substrins
''' sBefore       String before the string we are attempting to extract
''' sAfter        String after the string we are attempting to extract
''' sMidle        Optional String between Before and after
'''
'''Description:   Try to find the string within sLine located between
'''               sBefore and sAfter
'''
'''Note:          Case sensitive
'''
'''Example:       sExtract = LibStr_sExtractSubstring("BeforeExtractAfter", "Before", "After")
'''               Will return "Extract"
'''
'''Note:
'''
'''Warning:       If string sAfter is empty or not found this function will return everything
'''               following Before
'''
'''Updates Globals:
'''
'''Returns:       if Before Found
'''                  If after found
'''                     string between before and after
'''                  else
'''                     string following before
'''               elseIf after found
'''                     String preceding After
'''               elseIf Midle Found
'''                     Midle as return value
'''                     byref - Strings before and after the Midle
'''               else
'''                  empty string
'''
'''Programmer: Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_sExtractSubstring"

   Dim sResult             As String
   Dim lStrLen             As Long
   Dim lPosStart           As Long
   Dim lPosEnd             As Long
   
On Error GoTo ErrHnd

   '''If Before Found
   If sBefore <> "" Then
      '''If before is specified
      
      '''Find begining of the before string in the Line
      lPosStart = InStr(sLine, sBefore)
      '''If Before Found
      If lPosStart > 0 Then
         lStrLen = Len(sLine)
         '''Determine first position after sBefore
         lPosStart = lPosStart + Len(sBefore)
         '''If there are anything left after before string
         If lPosStart <= lStrLen Then
            '''If after is not specified
            If Not (sAfter = "") Then
               '''Determine position of the string sAfter
               lPosEnd = InStr(lPosStart, sLine, sAfter)
               '''If string after found
               If lPosEnd > 1 Then
                  '''If at least on character between before and after
                  If lPosEnd > lPosStart Then
                     '''Get everything in between
                     sResult = Mid$(sLine, lPosStart, lPosEnd - lPosStart)
                  End If
                  '''else return empty string
               Else
                  '''Get everything after before string
                  sResult = Mid$(sLine, lPosStart)
               End If
             Else
                '''Get everything after the string
                sResult = Mid$(sLine, lPosStart)
             End If
         End If
      End If
   ElseIf sAfter <> "" Then
      '''If after is found
      lPosStart = InStr(sLine, sAfter)
      '''If After Found
      If lPosStart > 0 Then
         '''Get everything before before string
         sResult = Left$(sLine, lPosStart - 1)
      End If
   ElseIf sMidle <> "" Then
      '''If midle is specified
      lPosStart = InStr(sLine, sMidle)
      '''If Midle found
      If lPosStart > 0 Then
         sResult = sMidle
         '''Get everything before and after the string
         sBefore = Left$(sLine, lPosStart - 1)
         sAfter = Mid$(sLine, lPosStart + Len(sMidle))
      End If
   End If

   '''Return result
   LibStr_sExtractSubstring = sResult
   
Exit Function
ErrHnd:
   LibStr_sExtractSubstring = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Public Function LibStr_sBuildSentence(ParamArray sWords()) As String
'''************************************************************************
'''Arguments:
''' sWords        Array of words to be combined into Sentence
'''
'''Description:   Process specified group of words
'''               Convert all words to lower Case
'''               Capitalize First Word and
'''               Remove accessed spaces
'''Note:
'''
'''Warning:       sWords must contain at least one word with at least one
'''               not space character
'''
'''Updates Globals:
'''Returns:       If successful
'''                  Sentence
'''               Else
'''                  empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_sBuildSentence"
   
   Dim sResult       As String
   Dim lIdx          As Long
   Dim sWord         As String
   
On Error GoTo ErrHnd
   
   For lIdx = LBound(sWords) To UBound(sWords)
      sWord = Trim$(sWords(lIdx))
      '''If all caps - use it as is
      If Not UCase$(sWord) = sWord Then
         sWord = LCase$(sWords(lIdx))
      End If
      '''If first not empty word is being added
      If (sResult = "") And (Not (sWord = "")) Then
         '''Capitalize first letter
         Mid$(sWord, 1, 1) = UCase$(Left$(sWord, 1))
         sResult = sWord
      Else
         '''Always capitalize ID
         If sWord = "id" Then
            sWord = UCase$(sWord)
         End If
         '''Omit space between word and (s)
         If sWord = "(s)" Then
            sResult = sResult & sWord
         Else
            sResult = sResult & " " & sWord
         End If
      End If
   Next lIdx
   
   '''Return result
   LibStr_sBuildSentence = sResult
   
Exit Function
ErrHnd:
   LibStr_sBuildSentence = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function


Public Function LibStr_sEndWith _
   (ByVal sSource As String, ByVal sEndWith As String, _
    ByVal bEndWith As Boolean, ByVal bTrim As Boolean, Optional bIgnoreCase As Boolean _
   ) As String
'''************************************************************************
'''Arguments:
''' sSource       Initial String Value
''' sEndWith      Character(s) to be evaluated
''' bEndWith      Request to remove or to add character(s) specified in the sEndWith
''' bTrim         Request to trim before processing
''' bIgnoreCase
'''   Indicates whether evaluation should be case sensitive
'''
'''Description:   Depending on value of bEndWith
'''               add or remove sEndWith from/to end of the specified string
'''
'''Example:       Build full path name of the file, Make sure Path ends with backslash
'''               sPath = LibStr_sEndWith(sPath, "\", true, true)  & sFileName
'''
'''Note:          Source string is passed byval - It will not be modified unless return value is assigned to it
'''               If end with string is empty - source will be returned unmodified
'''               If trim is requested - result will always be returned without leading or trailing spaces
'''               Trim applies to source string only
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       Processed specified string
'''               If error
'''                  empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_sEndWith"

   Dim sResult          As String
   Dim bDone            As Boolean
   Dim lMatchLen        As Long
   Dim sWith            As String
   
On Error GoTo ErrHnd
   
   
   If bTrim Then
      sSource = Trim$(sSource)
   End If
   sResult = sSource
   lMatchLen = Len(sEndWith)
   sWith = sEndWith
   
   '''If there is anything to look for
   If (lMatchLen > 0) Then
      If bIgnoreCase Then
         sEndWith = UCase$(sEndWith)
         sResult = UCase$(sResult)
      End If
      '''If specified pattern must be present
      If bEndWith Then
         '''If missing
         If Not Right$(sResult, lMatchLen) = sEndWith Then
            '''Add it
            sResult = sSource & sWith     '''Preserve case
         End If
      Else
         '''Specified pattern must not be present
         Do
            '''If pattern found
            If Right$(sResult, lMatchLen) = sEndWith Then
               '''Remove Specified pattern
               sResult = Left$(sResult, (Len(sResult) - lMatchLen))
            Else
               bDone = True
            End If
         Loop Until bDone
         If bIgnoreCase Then
            sResult = Left$(sSource, Len(sResult))
         End If
      End If
   End If
   
   '''Return result
   LibStr_sEndWith = sResult
   
Exit Function
ErrHnd:
   LibStr_sEndWith = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function
Public Function LibStr_sEmptyIfNull(ByVal vStr As Variant, Optional ByVal bTrim As Boolean) As String
'''************************************************************************
'''Arguments:
''' vStr          Value to be evaluated
''' bTrim         Optional request to trim return value
'''
'''Description:   Evaluate specified value and if it is not null and
'''               string return it else return empty string
'''
'''
'''Note:          If specified value is not string empty string will be returned
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       if specified value is string and not null
'''                  specified value
'''               else or error
'''                  empty string
'''
'''
'''Programmer: Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_sEmptyIfNull"

   Dim sResult             As String

On Error GoTo ErrHnd
   
   sResult = ""

   Select Case VarType(vStr)
      Case vbString
         sResult = vStr
   
      Case vbInteger To vbDate, vbBoolean, vbDecimal, vbByte, vbCurrency, vbDouble, vbSingle, vbInteger, vbLong
         sResult = Format$(vStr, "")
         
   End Select
   
   If bTrim Then
      sResult = Trim$(sResult)
   End If
   
   '''Return result
   LibStr_sEmptyIfNull = sResult
   
Exit Function
ErrHnd:
   LibStr_sEmptyIfNull = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function

Public Function LibStr_lReplaceSubString _
   (ByRef sSource As String, ByVal sStrToReplace As String, ByVal sReplaceStr As String, _
    ByVal bCaseSensitive As Boolean, Optional lMaxReplacements As Long = 1000 _
   ) As Long
'''************************************************************************
'''Arguments:
''' sSource       Source String to be evaluated and possibly modified
''' sStrToReplace Sub string of the sSource to be replaced with sReplaceStr
''' sReplaceStr   String to be used in the place of the sStrToReplace
''' bCaseSensitive Flag indicating whether search should be case sensitive
'''
'''Description:   Attempt to find sStrToReplace within sSource
'''               Ignore case if not bCaseSensitive
'''               If found replace all occurrences with sReplaceStr
'''
'''Note:          Case sensitive
'''
'''Example:       sExtract = LibStr_bReplaceSubString("ABC_23_HIJ", "123", "EFG", false)
'''               Will result in "ABC_EFG_HIJ"
'''
'''Note:          If sStrToReplace is not found sSource will not be changed and
'''               function will return true
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If no errors   - Number of replacements
'''               Else           - -1
'''
'''Programmer: Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME        As String = "LibStr_sExtractSubstring"

   Dim lResult             As Long
   Dim lLenSource          As Long
   Dim lLenStrToReplace    As Long
   Dim lPosStart           As Long
   Dim lPosEnd             As Long
   Dim sBefore             As String
   Dim sAfter              As String
   Dim bStrToStrSame       As Boolean
   Dim sSourceOrg          As String
   
On Error GoTo ErrHnd

   sSourceOrg = sSource
   lLenSource = Len(sSource)
   lLenStrToReplace = Len(sStrToReplace)
   If bCaseSensitive Then
      bStrToStrSame = (sStrToReplace = sReplaceStr)
   Else
      bStrToStrSame = UCase$(sStrToReplace) = UCase$(sReplaceStr)
   End If
   
   '''If source and string to replace are not empty string and replacement can change the source
   If (lLenSource > 0) And (lLenStrToReplace > 0) And (Not bStrToStrSame) Then
      lPosStart = 1
      Do
         If bCaseSensitive Then
            lPosStart = InStr(lPosStart, sSource, sStrToReplace)
         Else
            lPosStart = InStr(lPosStart, UCase$(sSource), UCase$(sStrToReplace))
         End If
         '''If found
         If lPosStart > 0 Then
            '''Store source contents before beginning of the string to replace
            sBefore = Left$(sSource, lPosStart - 1)
            '''Store source contents after end of the string to replace
            sAfter = Mid$(sSource, lPosStart + lLenStrToReplace)
            '''Replace substring
            sSource = sBefore & sReplaceStr & sAfter
            lPosStart = Len(sBefore & sReplaceStr) + 1
            lResult = lResult + 1
         End If
      Loop Until (lPosStart = 0) Or (lResult > lMaxReplacements)
   End If

   If (lResult > lMaxReplacements) Then
      lResult = -1
      sSource = sSourceOrg
   End If
   
   '''Return result
   LibStr_lReplaceSubString = lResult
   
Exit Function
ErrHnd:
   sSource = sSourceOrg
   LibStr_lReplaceSubString = -1
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
Exit Function
End Function


Public Function LibStr_sStartWith _
   (ByVal sSource As String, ByVal sStartWith As String, _
    ByVal bStartWith As Boolean, ByVal bTrim As Boolean, Optional bIgnoreCase As Boolean _
   ) As String
'''************************************************************************
'''Arguments:
''' sSource          Initial String Value
''' sStartWith   character(s) to be evaluated
''' bStartWith       Request to remove or to add character(s) specified in the sStartWith
''' bTrim            Request to trim before processing
''' bIgnoreCase
'''   Indicates whether evaluation should be case sensitive
'''
'''Description:      Depending on value of bStartWith
'''                  add or remove sStartWith from/to begining of the specified string
'''
'''Example:
'''
'''Note:             Source string is passed byval
'''                  - It will not be modified unless return value is assigned to it
'''                  If trim is requested
'''                     Result will always be returned without leading or trailing spaces
'''                  If start with string is empty - source will be returned unmodified
'''Warning:
'''
'''
'''Updates Globals:
'''
'''Returns:          Processed specified string
'''                  If error
'''                     empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_sStartWith"

   Dim sResult          As String
   Dim bDone            As Boolean
   Dim lMatchLen        As Long
   Dim sWith            As String
   
On Error GoTo ErrHnd
   
   If bTrim Then
      sSource = Trim$(sSource)
   End If
   sResult = sSource
   lMatchLen = Len(sStartWith)
   sWith = sStartWith
   
   '''If there is anything to look for
   If (lMatchLen > 0) Then
      If bIgnoreCase Then
         sStartWith = UCase$(sStartWith)
         sResult = UCase$(sResult)
      End If
      '''If specified pattern must be present
      If bStartWith Then
         '''If missing
         If Not Left$(sResult, lMatchLen) = sStartWith Then
            '''Add it
            sResult = sWith & sSource    '''Preserve case
         End If
      Else
         '''Specified pattern must not be present
         Do
            '''If pattern found
            If Left$(sResult, lMatchLen) = sStartWith Then
               '''Remove Specified pattern
               sResult = Right$(sResult, (Len(sResult) - lMatchLen))
            Else
               bDone = True
            End If
         Loop Until bDone
         If bIgnoreCase Then
            sResult = Right$(sSource, Len(sResult))
         End If
      End If
   Else
      sResult = sSource
   End If
   
   '''Return result
   LibStr_sStartWith = sResult
   
Exit Function
ErrHnd:
   LibStr_sStartWith = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Function
End Function


Public Function LibStr_sReplaceChar _
   (ByVal sSource As String, ByVal sCharSet As String, _
    ByRef bErr As Boolean, _
    Optional ByVal bSetContainsValidChars As Boolean, _
    Optional ByVal sReplaceChar As String = "" _
   ) As String
'''***********************************************************************
'''Arguments:
''' sSource       string to evaluated and possibly, modified
''' sCharSet      string containing set of validation characters
''' bErr          Indicates whether error occured during this procedure
''' bSetContainsValidChars
'''               Indicates whether sCharSet contains valid or invalid characters
'''               - defaults to false - invalid
''' sReplaceChar  Character to be used in place of invalid characters
'''               - defaults to empty string - remove invalid
'''
'''Description:   This procedure evaluates the passed string, if it
'''               contains the characters specified in the sCharSet
'''               string - remove them from the source.
'''
'''Note:          If sCharSet is an empty or a null string - source
'''               will be returned unmodified.
'''
'''Uses Modules:  None
'''
'''Uses Globals:  None
'''
'''Returns:       If no errors
'''                  Processed string with invalid characters removed or replaced
'''               Else
'''                  unmodified source string - see also bErr parameter
'''***********************************************************************
   Const sPROC_NAME        As String = msMODULE_NAME & "LibStr_sReplaceChar"
   
   Dim sResult             As String
   Dim lIdx                As Long

On Error GoTo ErrHnd

   If Not (sCharSet = "") Then
      If bSetContainsValidChars Then
         '''Keep if found
         For lIdx = 1 To Len(sSource)
            If InStr(sCharSet, Mid$(sSource, lIdx, 1)) > 0 Then
               sResult = sResult & Mid$(sSource, lIdx, 1)
            Else
               sResult = sResult & sReplaceChar
            End If
         Next
      Else
         '''Keep if not found
         For lIdx = 1 To Len(sSource)
            If InStr(sCharSet, Mid$(sSource, lIdx, 1)) = 0 Then
               sResult = sResult & Mid$(sSource, lIdx, 1)
            Else
               sResult = sResult & sReplaceChar
            End If
         Next
      End If
   Else
      sResult = sSource
   End If

   LibStr_sReplaceChar = sResult
   
Exit Function
ErrHnd:
   LibStr_sReplaceChar = sSource
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function


Public Function LibStr_bTxtEql(ByVal sStr1 As String, ByVal sStr2 As String) As Boolean
'''************************************************************************
'''Arguments:
''' sStr1
''' sStr2
'''
'''Description:   Compare specified strings ignoring presiding and trailing
'''               spaces and letter case
'''
'''Note:
'''
'''Warning:mpty string
'''
'''Updates Globals:
'''
'''Returns:       If contents(ignoring presiding and trailing spaces and case) of both strings are the same
'''                  True
'''               Else or error
'''                  False
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibStr_bTxtEql"

On Error GoTo ErrHnd
   
   '''Return result
   LibStr_bTxtEql = (UCase$(Trim$(sStr1)) = UCase$(Trim$(sStr2)))
   
Exit Function
ErrHnd:
   LibStr_bTxtEql = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Public Function LibStr_sReplaceCharToChar _
   (ByVal sSource As String, ByVal sCharSet As String, ByVal sReplaceSet As String, _
    ByRef bErr As Boolean) As String
'''***********************************************************************
'''Arguments:
''' sSource       string to evaluated and possibly, modified
''' sCharSet      string containing set of validation characters
''' sReplaceSet   Characters to be used in place of invalid characters
'''               must be equal length with sCharSet
''' bErr          Indicates whether error occured during this procedure
'''
'''Description:   This procedure evaluates the passed string, if it
'''               contains the characters specified in the sCharSet
'''               string - replace them with corresponding characters
'''
'''Note:          If sCharSet is an empty string - source will be returned unmodified.
'''
'''Uses Modules:  None
'''
'''Uses Globals:  None
'''
'''Returns:       If no errors
'''                  Processed string with invalid characters removed or replaced
'''               Else
'''                  unmodified source string - see also bErr parameter
'''***********************************************************************
   Const sPROC_NAME        As String = msMODULE_NAME & "LibStr_sReplaceCharToChar"
   
   Dim sResult             As String
   Dim lIdx                As Long
   Dim lPos                As Long

On Error GoTo ErrHnd

   If Not (sCharSet = "") Then
      If Len(sReplaceSet) = Len(sCharSet) Then
         For lIdx = 1 To Len(sSource)
            lPos = InStr(sCharSet, Mid$(sSource, lIdx, 1))
            If lPos = 0 Then
               sResult = sResult & Mid$(sSource, lIdx, 1)
            Else
               sResult = sResult & Mid$(sReplaceSet, lPos, 1)
            End If
         Next
      Else
         bErr = True
         sResult = sSource
         Call gobjLibApp.LibApp_Error _
         (1, LibStr_sBuildSentence(gsSET & "s", gsARE, gsNOT, gsEQUAL, gsLENGTH), _
          msMODULE_NAME, sPROC_NAME, "" _
         )
      End If
   Else
      sResult = sSource
   End If

   LibStr_sReplaceCharToChar = sResult
   
Exit Function
ErrHnd:
   LibStr_sReplaceCharToChar = sSource
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Public Property Get psNUMBERS() As String
   psNUMBERS = "0123456789"
End Property


