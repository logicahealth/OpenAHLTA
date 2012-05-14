VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SyncAnP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class is used to make sure that the A/P Encounter RTF matches that data
'loaded by A/P.

Option Explicit

Private Enum enumLineType
  typDiag
  typOrder
End Enum

Private mDocLines  As Variant        'Array of lines from the plain text version of A/P Encounter RTF
Private mCurrentLine As Long         'Current line in the mDocLines array that is being verified
Private mLineType As enumLineType    'Type of line that is current being verified

'JT BEGIN SCRs 38958 to 38960 - removes comments and code annotations from the RTF so they don't affect validation
Private Function CleanRTF(ByVal RTF As String) As String

On Error GoTo ErrHandler

Dim sRTFOut As String
Dim lCurPos As Long
Dim lEndPos As Long
Dim sReplaceText As String

   sRTFOut = RTF
   
   'Find a comment/code tag
   lCurPos = InStr(1, sRTFOut, RTF_COMMENT_BEGIN_TAG)
   
   Do While lCurPos > 0
   
      'Find the end tag
      lEndPos = InStr(lCurPos, sRTFOut, RTF_COMMENT_END_TAG)
      
      If lEndPos > 0 Then
         'Get the string to replace and remove it from the RTF
         sReplaceText = Mid$(sRTFOut, lCurPos, (lEndPos + Len(RTF_COMMENT_END_TAG)) - lCurPos)
         Debug.Print "Replacing:" & vbCrLf & sReplaceText
         sRTFOut = Replace(sRTFOut, sReplaceText, "", 1)
      End If
      
      'Look for more comment/code tags
      lCurPos = InStr(lCurPos, sRTFOut, "{\b}{")
   Loop
   
   CleanRTF = sRTFOut
   
   Exit Function
   
ErrHandler:
  If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.CleanRTF", "A/P", vbExclamation

End Function
'JT END SCRs 38958 to 38960

'Removes any text in front of the diagnosis description or order name.
Private Function CleanLine() As String
Dim DiagnosisPriority As Integer
Dim CurrentLine As String
Dim DiagnosisLine As String

  On Error GoTo ErrHandler
  
  DiagnosisPriority = Val(mDocLines(mCurrentLine))   'Diagnosis have number in front of them
  
  DiagnosisLine = DiagnosisPriority & ". "
  If Left$(mDocLines(mCurrentLine), Len(DiagnosisLine)) <> DiagnosisLine Then DiagnosisPriority = 0
    
  If DiagnosisPriority > 0 Then
    'Diagnosis
    mLineType = typDiag
    CurrentLine = Trim$(Mid$(mDocLines(mCurrentLine), Len(CStr(DiagnosisPriority) & ". ")))
  Else
    'Order
    mLineType = typOrder
    CurrentLine = mDocLines(mCurrentLine)
    
    'Remove Consult Header if found
    CurrentLine = RemoveSubString("CONSULT(S):", CurrentLine)
      
    'Remove Procedure Header if found
    CurrentLine = RemoveSubString("PROCEDURE(S):", CurrentLine)
  
    'Remove Lab Header if found
    CurrentLine = RemoveSubString("LABORATORY:", CurrentLine)
  
    'Remove Med Header if found
    CurrentLine = RemoveSubString("MEDICATION(S):", CurrentLine)
  
    'Remove Rad Header if found
    CurrentLine = RemoveSubString("RADIOLOGY:", CurrentLine)
  
    'Remove PI Header if found
    CurrentLine = RemoveSubString("PATIENT INSTRUCTION(S):", CurrentLine)
  
    'Lastest Order Format has a dash at the beginning of an order
    If Left$(mDocLines(mCurrentLine), 1) = "-" Then
      CurrentLine = Trim$(Mid$(CurrentLine, 2))
    End If
    
    'Lastest Consult Format has Referred To at the beginning of an order
    If Left$(Trim$(UCase$(CurrentLine)), 12) = "REFERRED TO:" Then
      CurrentLine = Trim$(Mid$(CurrentLine, 13))
    End If
    
    CurrentLine = Replace(CurrentLine, vbTab, vbNullString)
  End If
  
  CleanLine = CurrentLine
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.CleanLine", "A/P", vbExclamation
  
End Function

'Checks to see that all diagnsoses in the A/P data are found in the RTF.
Private Function IsAnPDiagnosesValid(ByRef RTF As String) As Boolean
Dim Diagnosis As AnPDiagnosis
Dim MatchFound As Boolean
Dim Convert As IConvert
Dim AllAltDescriptions As Collection
Dim i As Integer
Dim AllDiagnoses As AnPDiagnoses

  On Error GoTo ErrHandler
  
  Set AllDiagnoses = gBB.AllDiagnoses
  
  If AllDiagnoses.Count = 0 Then
    MatchFound = True
  End If
  
  Set Convert = New IConvert
  
  For Each Diagnosis In AllDiagnoses.Filter
    
    Set AllAltDescriptions = Convert.AnP_PossibleSnoDescriptions(Diagnosis.SnoID)
    
    For i = 1 To AllAltDescriptions.Count
      MatchFound = InStr(1, UCase$(RTF), UCase$(AllAltDescriptions.Item(i))) > 0
      If MatchFound Then Exit For
    Next
    
    If Not MatchFound Then Exit For
  Next
  
  IsAnPDiagnosesValid = MatchFound
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.IsAnPDiagnosesValid", "A/P", vbExclamation

End Function

'Checks to see that all orders in the A/P data are found in the RTF.
Private Function IsAnPOrdersValid(ByRef RTF As String) As Boolean
Dim Order As iOrder
Dim MatchFound As Boolean
Dim Orders As Collection
Dim i As Long
'Dim Consult As OrderConsult

  On Error GoTo ErrHandler
  
  If gBB.AllOrders.Count = 0 Then
    MatchFound = True
  End If
  
  Set Orders = gBB.AllOrders.Filter
  
  For i = 1 To Orders.Count
    Set Order = Orders.Item(i)
    
    MatchFound = InStr(1, UCase$(RTF), UCase$(Order.Name)) > 0
  
'    If Not MatchFound Then
'      If TypeName(Order) = "OrderConsult" Then
'        Set Consult = Order
'        'JT SCR 44366 - during the comparison use the same RTFCodeFilter function that was used to build the RTF
'        MatchFound = InStr(1, UCase$(RTF), UCase$(RTFCodeFilter(Consult.ReferTo))) > 0
'      End If
'    End If
        
    If Not MatchFound Then Exit For
  Next
  
  IsAnPOrdersValid = MatchFound

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.IsAnPOrdersValid", "A/P", vbExclamation

End Function

'Checks the collection of diagnoses to see if it can find a match for the String line passed in.
Private Function IsLineDiagnosisValid(ByVal Line As String) As Boolean
Dim Diagnoses As Collection
Dim Diagnosis As AnPDiagnosis
Dim MatchFound As Boolean
Dim ColonLocation As Long

  On Error GoTo ErrHandler
  
  Set Diagnoses = gBB.AllDiagnoses.Filter
  
  For Each Diagnosis In Diagnoses
    'SCR 27893 Phil Crowder 11/11/02
    MatchFound = InStr(1, UCase$(Diagnosis.FullName), UCase$(Line)) > 0
    
    'SCR 28663  3/11/02  Phil Crowder  Changed RTF format of Diagnosis RTF
    If Not MatchFound Then
      ColonLocation = InStr(1, Line, ":")
      
      If ColonLocation > 0 Then
        Line = Mid$(Line, 1, ColonLocation - 1)
        
        MatchFound = InStr(1, UCase$(Diagnosis.FullName), UCase$(Line)) > 0
      End If
    End If
    
    If MatchFound Then Exit For
  Next
    
  IsLineDiagnosisValid = MatchFound
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.IsLineDiagnosisValid", "A/P", vbExclamation
  
End Function

'Checks the collection of orders to see if it can find a match for the String line passed in.
Private Function IsLineOrderValid(ByRef Line As String) As Boolean
Dim Order As iOrder
Dim MatchFound As Boolean
Dim Orders As Collection
Dim i As Long
'Dim Consult As OrderConsult

  On Error GoTo ErrHandler
  
  Set Orders = gBB.AllOrders.Filter
  
  For i = 1 To Orders.Count
    Set Order = Orders.Item(i)
    
    MatchFound = InStr(1, Trim$(UCase$(Line)), Trim$(UCase$(Order.Name))) > 0
  
'    If Not MatchFound Then
'      If TypeName(Order) = "OrderConsult" Then
'        Set Consult = Order
'        MatchFound = InStr(1, Trim$(UCase$(Line)), Trim$(UCase$(Consult.ReferTo))) > 0
'      End If
'    End If
        
    If MatchFound Then Exit For
  Next
  
  IsLineOrderValid = MatchFound
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.IsLineOrderValid", "A/P", vbExclamation
  
End Function

'Checks current line to see if it breaks any of these rules.  If it does then we want to skip
'this line.
Private Function IsLineValid() As Boolean
Dim CurrentLine As String

  On Error GoTo ErrHandler
  
  'Remove tabs from line
  CurrentLine = Replace(mDocLines(mCurrentLine), vbTab, vbNullString)
  
  'Skip Blank lines
  If Len(Trim$(CurrentLine)) = 0 Then Exit Function
  
  'Skip Comment lines
  If InStr(1, UCase$(Trim$(CurrentLine)), "COMMENT") Then Exit Function
  
  'Skip send order line if exists
  If Left$(UCase$(Trim$(CurrentLine)), 10) = "ORDER DATE" Then Exit Function
  
  'Skip Clinical Impression line if exists
  If Left$(UCase$(Trim$(CurrentLine)), 19) = "CLINICAL IMPRESSION" Then Exit Function

  'Skip Consult Header line if exists
  If UCase$(Trim$(CurrentLine)) = "CONSULT(S):" Then Exit Function

  'Skip Lab Header line if exists
  If UCase$(Trim$(CurrentLine)) = "LABORATORY:" Then Exit Function

  'Skip Lab Header line if exists
  If UCase$(Trim$(CurrentLine)) = "LABORATORY(IES):" Then Exit Function

  'Skip Med Header line if exists
  If UCase$(Trim$(CurrentLine)) = "MEDICATION(S):" Then Exit Function

  'Skip PI Header line if exists
  If UCase$(Trim$(CurrentLine)) = vbTab & "PATIENT INSTRUCTION(S):" Then Exit Function

  'Skip Procedure Header line if exists
  If UCase$(Trim$(CurrentLine)) = "PROCEDURE(S):" Then Exit Function

  'Skip Rad Header line if exists
  If UCase$(Trim$(CurrentLine)) = "RADIOLOGY:" Then Exit Function

  'Skip Rad Header line if exists
  If UCase$(Trim$(CurrentLine)) = "RADIOLOGY(IES):" Then Exit Function

  'Skip Unassociated Orders Header line if exists
  If UCase$(Trim$(CurrentLine)) = "--> UNASSOCIATED ORDERS AND PROCEDURES <--" Then Exit Function

  IsLineValid = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.IsLineValid", "A/P", vbExclamation
  
End Function

'Checks that all of the diagnoses and orders in the RTF are in A/P's version of the data.
Private Function IsRTFDiagnosesOrdersValid(ByRef RTF As String) As Boolean
Dim ValidLine As Boolean
Dim PlainText As String
Dim CurrentLine As String

  On Error GoTo ErrHandler
  
  'JT SCRs 38958 to 38960 - Strip out any comments that may be in the rtf so they don't affect validation
  PlainText = TurnRTFtoPlainText(CleanRTF(RTF))
  mDocLines = Split(PlainText, vbLf)
  
  If UBound(mDocLines) < 2 Then
    IsRTFDiagnosesOrdersValid = True
    Exit Function
  End If
  
  mCurrentLine = 0
  CurrentLine = mDocLines(mCurrentLine)
  'SCR 107019
  If Len(CurrentLine) = 0 Then
      mCurrentLine = 1
      CurrentLine = mDocLines(mCurrentLine)
  End If
  
  Do While Len(CurrentLine) > 0
    CurrentLine = NextLine
    If Len(CurrentLine) = 0 Then Exit Do
    
    If mLineType = typDiag Then
      ValidLine = IsLineDiagnosisValid(CurrentLine)
    Else
      ValidLine = IsLineOrderValid(CurrentLine)
    End If
    
    If Not ValidLine Then Exit Do
  Loop

  IsRTFDiagnosesOrdersValid = ValidLine
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.IsRTFDiagnosesOrdersValid", "A/P", vbExclamation
  
End Function

'Gets the next line with a diagnosis description or order name.
Private Function NextLine() As String
Dim ValidLine As Boolean
    
  On Error GoTo ErrHandler
  
  Do While Not ValidLine
    mCurrentLine = mCurrentLine + 1
    If mCurrentLine > UBound(mDocLines) Then Exit Do
    
    ValidLine = IsLineValid
  Loop
      
  If mCurrentLine > UBound(mDocLines) Then
    NextLine = vbNullString
  Else
    NextLine = CleanLine
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.NextLine", "A/P", vbExclamation
  
End Function

'Removes a sub-string value from the String Line passed in.
'Returns:
' String without sub string
Private Function RemoveSubString(ByRef SubString As String, ByRef Line As String) As String
Dim HeaderBegins As Integer

  On Error GoTo ErrHandler
  
  HeaderBegins = InStr(1, UCase$(Line), UCase$(SubString))
  
  If HeaderBegins > 0 Then
    RemoveSubString = Trim$(Mid$(Trim$(Line), Len(SubString) + 1 + HeaderBegins))
  Else
    RemoveSubString = Line
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.RemoveSubString", "A/P", vbExclamation

End Function

'JT SCR 44681 - This method checks to ensure that all the orders referenced in a diagnosis order association
'are present in the appropriate orders collect.  This could be the result of a failed save operation that the
'current user may not be aware of.  If the condition exists, a warning will be presented to the user instructing
'them to close AnP, print the encounter document, and reconstruct the information in A/P
Public Function SyncOrderWarningRequired() As Boolean

On Error GoTo ErrHandler

Dim sWarningMsg As String
Dim objDx As AnPDiagnosis
Dim objAssoc As OrderAssociation

   For Each objDx In gBB.AllDiagnoses.Filter(fltActive)
      For Each objAssoc In objDx.Filter(typConsult, fltActive)
         If gBB.AllOrders.Item(typConsult, objAssoc.OrderKey) Is Nothing Then
            'Consult order was not found - Add item to warning list and remove for association collection
            sWarningMsg = sWarningMsg & vbCrLf & vbCrLf & objDx.Priority & "." & vbTab & objDx.FullICDCode & " - " & objDx.FullName & vbCrLf & _
               vbTab & "Unidentifiable Consult Order Association for DataID = " & objAssoc.OrderKey
            objDx.RemoveOrderAssocation typConsult, objAssoc.OrderKey
         End If
      Next
   Next
   
   If sWarningMsg <> vbNullString Then
      'Build and show the final message
      sWarningMsg = "The Encounter you have selected contains an association(s) to an " & _
            "order that can no longer be located.  Please make a note of the information below and " & _
            "contact the MHS Help Desk for support. " & sWarningMsg & vbCrLf & vbCrLf & _
            "Note:  If you close A/P without making changes, the SF600 may still reflect the missing order(s) " & _
            "and aid in identifying the missing data."
      WarningDialog.Display sWarningMsg, True
   End If

ErrHandler:
   'Clean Up Objects
   Set objDx = Nothing
   Set objAssoc = Nothing
   
   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.SyncOrderWarningRequired", "A/P", vbExclamation
  
End Function
'JT END SCR 44681

'SCR 16771 Phil Crowder 10/24/01
'SCR 22949 Phil Crowder 04/13/02
'SCR 22899 Phil Crowder 04/13/02
'This method is needed because before build 535 of A/P there is a possibilty that
'the encounter RTF does not match what was saved in A/P.  This function is needed
'to check older encounter if they are amended.
'This method checks to make sure that the diagnoses and related items that are in A/P
'match the Encounter RTF.
'Parameters:
' RTF   Encounter RTF (for A/P section only)
'Returns:
' True if a waring is required and false if everything matches
Public Function SyncWarningRequired(ByRef RTF As String) As Boolean
Dim Valid As Boolean

  On Error GoTo ErrHandler
  
  Valid = IsAnPDiagnosesValid(RTF)
  
  If Valid Then
    Valid = IsAnPOrdersValid(RTF)
     
    If Valid Then
      Valid = IsRTFDiagnosesOrdersValid(RTF)
    End If
  End If

  SyncWarningRequired = Not Valid
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.SyncWarningRequired", "A/P", vbExclamation
  
End Function

'Uses RTF control on SummaryDialog to get plain text version of the A/P RTF.
Private Function TurnRTFtoPlainText(ByRef RTF As String) As String
Dim Summary As SummaryDialog

  On Error GoTo ErrHandler
  
  Set Summary = New SummaryDialog
  
  Summary.InitRTF RTF
  TurnRTFtoPlainText = Summary.rtfPreviewSummary.Text
  
  Unload Summary
  Set Summary = Nothing
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "SyncAnP.TurnRTFtoPlainText", "A/P", vbExclamation
  
End Function
