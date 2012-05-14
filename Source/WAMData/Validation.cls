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
Attribute VB_Name = "Validation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Errors"
Attribute VB_Ext_KEY = "Member1" ,"Lookups"
Option Explicit
Option Compare Text

Public ValidationMode As EnumValidationMode

Private mvarErrors As Errors
Private mvarICD9s As ICD9s
Private mvarCPT4s As CPT4s
Private mvarEMs As EandMs
Private mvarLookups As Lookups

Const DECODE_ERROR = "ERROR - Code not found"

Public Property Get Lookups() As Lookups
    Set Lookups = mvarLookups
End Property
Public Property Set Lookups(vData As Lookups)
    Set mvarLookups = vData
End Property

Friend Sub Init(ICD9s As ICD9s, CPT4s As CPT4s, EMs As EandMs)
    
    Set mvarICD9s = ICD9s
    Set mvarCPT4s = CPT4s
    Set mvarEMs = EMs
    
End Sub
Friend Sub Release()

    Set mvarICD9s = Nothing
    Set mvarCPT4s = Nothing
    Set mvarEMs = Nothing

End Sub

Public Property Get Errors() As Errors
    Set Errors = mvarErrors
End Property
Public Property Set Errors(vData As Errors)
    Set mvarErrors = vData
End Property



Friend Function GetDescription(ByVal Code As String, ByVal Lookup As EnumLookUps) As String
    
    
    Dim sz As String
    Dim rs As adodb.Recordset
    
    On Error GoTo errXIT
    
    'During Persistence, there is no need to look up descriptions
    'wastes bandwidth during transfer AND
    'risks sending non-soap compatible characters
    If ValidationMode = vmNone Then Exit Function
    
    'Msgbox "Lookup Code Goes Here"
    Select Case Lookup
    'Requires Lookup against Tables
        Case elApptType
            Set rs = Lookups.Item(Lookup).Recordset
            rs.MoveFirst
            rs.Find "Code LIKE '" & Code & "*'", , adSearchForward, adBookmarkFirst
            If Not rs.EOF Then
                sz = rs.Fields("DESCRIPTION").Value
            Else
                sz = DECODE_ERROR
            End If
            
        Case elEandM, elCPT4, elICD9, elMEPRS, elApptType
            Set rs = Lookups.Item(Lookup).Recordset
            rs.MoveFirst
            rs.Find "Code='" & Code & "'", , adSearchForward, adBookmarkFirst
            If Not rs.EOF Then
                sz = rs.Fields("DESCRIPTION").Value
            Else
                sz = DECODE_ERROR
            End If
            
        Case elModifiersEandM, elModifiersCPT4
            Set rs = Lookups.Item(CStr(Lookup)).Recordset
            rs.MoveFirst
            rs.Find "Code='" & Code & "'", , adSearchForward, adBookmarkFirst
            If Not rs.EOF Then
                sz = rs.Fields("SHORT_DESCRIPTION").Value
            Else
                sz = DECODE_ERROR
            End If
            
        Case elProviderIEN
            'No Provider Validations - File too big
'            Set rs = Lookups.Item(CStr(elProviderIEN)).Recordset
'            rs.MoveFirst
'            rs.Find "IEN='" & Code & "'", , adSearchForward, adBookmarkFirst
'            If Not rs.EOF Then
'                sz = rs.Fields("Name").Value
'            Else
'                sz = DECODE_ERROR
'            End If
            
    'Simple Lookups
        Case elAdminCode
            Select Case Code
                Case 1: sz = "Consultation Requested"
                Case 2: sz = "Referred to Another Provider"
                Case 3: sz = "Convalescent leave"
                Case 4: sz = "Medical Board"
                Case 5: sz = "Medical Hold"
                Case Else: sz = DECODE_ERROR
            End Select
        Case elDisposition
            Select Case Code
                Case "95": sz = "Admitted"
                Case "96": sz = "Expired"
                Case "90": sz = "Immediate Referral"
                Case "06": sz = "Left Against Medical Advice"
                Case "97": sz = "Left without being seen"
                Case "85": sz = "Released w/ work duty limitations"
                Case "91": sz = "Released w/o limitations"
                Case "05": sz = "Sick at home/quarters"
                Case Else: sz = DECODE_ERROR
            End Select
        Case elProviderRole
            Select Case Code
                Case "1": sz = "Attending"
                Case "2": sz = "Assisting"
                Case "3": sz = "Supervising"
                Case "4": sz = "Nurse"
                Case "5": sz = "Para-professional"
                Case Else: sz = DECODE_ERROR
            End Select
    End Select
    
    GetDescription = sz
    
errXIT:
'    If Code <> "" Then
'        If sz = DECODE_ERROR Then
'                Debug.Assert False
'        End If
'    End If
End Function

Public Function IsValidCode(ByVal Code As String, ByVal TypeOfCode As EnumLookUps) As Boolean
Dim Description As String

    On Error GoTo errXIT

    Description = GetDescription(Code, TypeOfCode)
    
    IsValidCode = Description <> DECODE_ERROR
    
errXIT:

End Function

Friend Sub ValidateCode(ByVal ObjectName As String, ByVal PropName As String, _
                             ByVal Code As String, ByVal Lookup As EnumLookUps)
    
    Dim sz As String
    Dim szErr As String
    Dim szErr1 As String
    
    If ValidationMode = vmNone Then Exit Sub
    
    '1 Find Code in Ref Table
    sz = GetDescription(Code, Lookup)
    
    '2 If Not Found raise Error
    If sz = DECODE_ERROR Then
    
        szErr = "'%2' is not a valid value for %4.%1. '%2' cannot be found in the '%3' code reference."
        szErr = Replace(szErr, "%1", PropName)
        szErr = Replace(szErr, "%2", Code)
        szErr = Replace(szErr, "%4", ObjectName)
        
        szErr1 = "'%2' is not a valid value for %1. '%2' cannot be found in the '%3' code reference."
        szErr1 = Replace(szErr1, "%1", PropName)
        szErr1 = Replace(szErr1, "%2", Code)
        szErr1 = Replace(szErr1, "_", " ")
        
        Select Case Lookup
            Case elEandM
                szErr = Replace(szErr, "%3", "E&M")
                szErr1 = Replace(szErr1, "%3", "E&M")
            Case elCPT4
                szErr = Replace(szErr, "%3", "CPT4")
                szErr1 = Replace(szErr1, "%3", "CPT4")
            Case elICD9
                szErr = Replace(szErr, "%3", "ICD9")
                szErr1 = Replace(szErr1, "%3", "ICD9")
            Case elModifiersCPT4
                szErr = Replace(szErr, "%3", "CPT4 Modifiers")
                szErr1 = ""
            Case elModifiersEandM
                szErr = Replace(szErr, "%3", "E&M Modifiers")
                szErr1 = ""
            Case elAdminCode
                szErr = Replace(szErr, "%3", "Admin")
                szErr1 = ""
            Case elDisposition
                szErr = Replace(szErr, "%3", "Disposition")
                szErr1 = ""
            Case elProviderIEN
                szErr = Replace(szErr, "%3", "Provider IEN")
            Case elProviderRole
                szErr1 = ""
                szErr = Replace(szErr, "%3", "Provider Role")
            Case elApptType
                szErr = Replace(szErr, "%3", "Appointment Type")
                szErr1 = ""
        End Select
        
        ProcessError vbObjectError + 1000, ObjectName & "." & PropName, szErr, szErr1

    
    End If
    
End Sub

Friend Sub ValidateColByMaxCount( _
        ByVal ObjectName As String, ByVal MethodName As String, _
        ByRef col As Collection, ByVal MaxCount As Long, ByVal Phrase As String)
        
    If ValidationMode = vmNone Then Exit Sub
    
    If col.Count >= MaxCount Then

        ProcessError vbObjectError + 1001, ObjectName & "." & MethodName, _
        "Only " & MaxCount & " " & Phrase & " are allowed."
        
    End If
    
End Sub

Friend Sub ValidatePropByList( _
        ByVal ObjectName As String, ByVal PropName As String, _
        ByRef NewValue As Variant, ParamArray ValueList() As Variant)
        
    Dim i As Long
    Dim szErr As String
    
    If ValidationMode = vmNone Then Exit Sub
    
    For i = LBound(ValueList) To UBound(ValueList)
        If NewValue = ValueList(i) Then
            Exit Sub
        End If
    Next
    
    szErr = "'%2' is not a valid value for %4-%1. Valid values must be in %3."
    szErr = Replace(szErr, "%1", PropName)
    szErr = Replace(szErr, "%2", NewValue)
    szErr = Replace(szErr, "%3", Join(ValueList, " or "))
    szErr = Replace(szErr, "%4", ObjectName)
    
    ProcessError vbObjectError + 1002, ObjectName & "." & PropName, szErr
'    ProcessError vbObjectError + 1002, ObjectName & "." & PropName, _
'    PropName & " must be in " & Join(ValueList, " or ") '
        
End Sub

Friend Sub ValidatePropByLength( _
        ByVal ObjectName As String, ByVal PropName As String, _
        ByRef NewValue As Variant, ByVal Length As Long)
        
    If ValidationMode = vmNone Then Exit Sub
    
    If Len(NewValue) > Length Then

        ProcessError vbObjectError + 1003, ObjectName & "." & PropName, _
        PropName & " cannot exceed " & Length & " chars."
        
    End If
    
End Sub

Friend Sub ValidatePropByIntegerRange( _
        ByVal ObjectName As String, ByVal PropName As String, _
        ByRef NewValue As Variant, ByVal LowRange As Long, ByVal HighRange As Long)
        
    If ValidationMode = vmNone Then Exit Sub
    
    If IsNumeric(NewValue) Then
        If Val(NewValue) >= LowRange And Val(NewValue) <= HighRange Then
            Exit Sub
        End If
    End If
    
    ProcessError vbObjectError + 1004, ObjectName & "." & PropName, _
    PropName & " must be a number from " & LowRange & " and " & HighRange & "."
    
End Sub
Friend Sub ValidateDIAGNOSIS_PRIORITY(ByVal ObjectName As String, ByVal PropName As String, ByVal NewValue As String)
    'Format: '1234'
    'Required (at least one diagnosis pointer required per E&M code)
    'Only Allow chars 1,2,3,4. 1,2,3,4 are only allowed to be used once.
    'Only can use value if Dx_Priority is in Dx collection - ie Dx_Priorities are 1,2,3,
    'then value can be some combination of 1,2,3. Same as E&M validation
        
    Dim objICD As ICD9
    Dim col As New Collection
    Dim i As Long
    Dim v As Variant
    Dim bFound As Boolean
    Dim szICD As String
    
    If ValidationMode = vmNone Then Exit Sub
    
    If Len(NewValue) = 0 Then
        ProcessError vbObjectError + 10015, ObjectName & "." & PropName, "One or more E&M/CPT codes have not been associated with a diagnosis.", "One or more E&M codes have not been associated with a diagnosis."
        Exit Sub
    End If
    
    ValidateRequiredValue ObjectName, PropName, NewValue
    
    'Build Code
    For Each objICD In mvarICD9s
        Select Case objICD.ICD_DIAGNOSIS_PRIORITY
        Case "1" To "4"
            szICD = szICD & objICD.ICD_DIAGNOSIS_PRIORITY
        Case Else
        
        End Select
    Next
    
    'Must Have ICD9s before you can add CPT4s or E&Ms
    If Len(szICD) = 0 Then
        ProcessError vbObjectError + 1005, ObjectName & "." & PropName, _
        PropName & " cannot be set until ICD9s are added."
        Exit Sub
    End If
    
    'Check That Each Code is in ICD_DIAGNOSIS_PRIORITYs
    For i = 1 To Len(NewValue)
        If InStr(1, szICD, Mid(NewValue, i, 1)) = 0 Then
            ProcessError vbObjectError + 1005, ObjectName & "." & PropName, _
            PropName & ": " & NewValue & " does not conform to ICD_DIAGNOSIS_PRIORITYs: " & szICD & "."
            Exit Sub
        End If
    Next
    
    'Check For Duplicate Chars
    Set col = New Collection
    On Error GoTo errDups
    For i = 1 To Len(NewValue)
        col.Add "X", Mid(NewValue, i, 1)
    Next
    
    Exit Sub
    
errDups: 'Duplicate Chars
    ProcessError vbObjectError + 1006, ObjectName & "." & PropName, _
    PropName & ": " & NewValue & " cannot have characters."
    
End Sub
Friend Sub ValidateRequiredValue( _
        ByVal ObjectName As String, ByVal PropName As String, _
        ByRef NewValue As Variant)
        
        If IsDate(NewValue) Then
            If CDbl(NewValue) = 0 Then GoTo errXIT
        End If
        
        If Len(NewValue) > 0 Then Exit Sub
        
errXIT:
        ProcessError vbObjectError + 1010, ObjectName & "." & PropName, _
        PropName & " is a required value."

End Sub

Friend Sub ValidateDuplicateCode( _
        ByVal ObjectName As String, ByVal PropName As String, _
        ByVal NewValue As Variant)
        
        ProcessError vbObjectError + 1020, ObjectName & "." & PropName, _
        PropName & ": " & NewValue & " is a duplicate code. Duplicate codes are not allowed."

End Sub

Private Sub ProcessError(ByVal ErrNumber As Long, ByVal Source As String, ByVal Description As String, Optional ByVal UserDescription As String = vbNullString)
    
    Select Case ValidationMode
    Case vmNone
        'Do Nothing
    Case vmSaveToCollection
        Me.Errors.Add ErrNumber, Source, Description, UserDescription
        
    Case vmImmediate
        Err.Raise ErrNumber, Source, Description
        
    End Select
End Sub


Private Sub Class_Initialize()
    ValidationMode = vmNone
    Set mvarErrors = New Errors
    Set mvarLookups = New Lookups
End Sub
Private Sub Class_Terminate()
  Set mvarLookups = Nothing
  Set mvarLookups = Nothing
  Set mvarErrors = Nothing
   Set mvarErrors = Nothing
End Sub


