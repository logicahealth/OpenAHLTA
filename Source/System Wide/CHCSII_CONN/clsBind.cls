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
Attribute VB_Name = "clsBind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim moStatement As SqlServStatement
Dim mlInterval As Long
Public Function ProcessMatch(Match As String, Position As Long, Source As String) As String
Attribute ProcessMatch.VB_UserMemId = 0
    
'This function is called by the regular expression object so we can process the match.

On Error GoTo ErrHnd

Const cIgnoreChars As String = " " & vbCr & vbLf & vbCrLf & vbTab
Const cCompoundChars As String = "><!"
Dim lPos As Long
Dim sChar As String
Dim bFound As Boolean

    mlInterval = mlInterval + 1

    'Only add bind variables if the operator is an = sign
    'Search backwards for a "="
    lPos = Position + 1
    Do
        lPos = lPos - 1
        sChar = Mid(Source, lPos, 1)
    Loop Until sChar = "=" Or InStr(cIgnoreChars, sChar) = 0 Or lPos = 1
        
    'Check one more character back incase it found a ">=" or a "<="
    If sChar = "=" And lPos <> 1 Then
        sChar = Mid(Source, lPos - 1, 1)
        If InStr(cCompoundChars, sChar) > 0 Then
            'Found a compound operator, do nothing
        Else
            'Aleady had a legit "="
            sChar = "="
        End If
    End If
        
    If sChar <> "=" Then
        'Search forwards for a "="
        lPos = Position + Len(Match)
        Do
            lPos = lPos + 1
            sChar = Mid(Source, lPos, 1)
        Loop Until sChar = "=" Or InStr(cIgnoreChars, sChar) = 0 Or lPos > Len(Source)
    End If
    
    If sChar = "=" Then
        Dim oBVar As SqlServBindVar
        
        'Convert it to a bind variable
        Set oBVar = moStatement.BindVariables.Add
        oBVar.Value.InfoType = "textString"
        
        'Trim off any leading and trailing single quotes
        If Left(Match, 1) = "'" And Right(Match, 1) = "'" Then
            'The match is a string (expression)
            Match = Mid(Match, 2, Len(Match) - 2)
            'SF 3-17-04 Added a space before and after the bind variable to compensate for bad incoming sql
            ProcessMatch = " :X" & mlInterval & " "
        Else
            'Match is a number so explicitly convert it to a number
            'SF 3-17-04 Added a space before and after the bind variable to compensate for bad incoming sql
            ProcessMatch = " TO_NUMBER(:X" & mlInterval & ") "
        End If
        
        'Convert any chr(0) to '
        oBVar.Value.TextString = Replace(Match, Chr(0), "'")
        
        Set oBVar = Nothing
    Else
        ProcessMatch = Match
    End If
    
Exit Function

ErrHnd:
    'Raise the error
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Property Get Statement() As SqlServStatement

    Set Statement = moStatement
    
End Property

Public Property Set Statement(vData As SqlServStatement)

    Set moStatement = vData
    mlInterval = 0
    
End Property

Private Sub Class_Terminate()

    Set moStatement = Nothing
    
End Sub


