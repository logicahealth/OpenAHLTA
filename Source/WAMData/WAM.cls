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
Attribute VB_Name = "WAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Top level data object"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Appointment"
Attribute VB_Ext_KEY = "Member1" ,"Consults"
Attribute VB_Ext_KEY = "Member2" ,"CPT4s"
Attribute VB_Ext_KEY = "Member3" ,"CPT4s"
Attribute VB_Ext_KEY = "Member4" ,"Disposition"
Attribute VB_Ext_KEY = "Member5" ,"Disposition"
Attribute VB_Ext_KEY = "Member6" ,"EandMs"
Attribute VB_Ext_KEY = "Member7" ,"EandMs"
Attribute VB_Ext_KEY = "Member8" ,"Encounter"
Attribute VB_Ext_KEY = "Member9" ,"Encounter"
Attribute VB_Ext_KEY = "Member10" ,"ICD9s"
Attribute VB_Ext_KEY = "Member11" ,"ICD9s"
Attribute VB_Ext_KEY = "Member12" ,"Labs"
Attribute VB_Ext_KEY = "Member13" ,"Meds"
Attribute VB_Ext_KEY = "Member14" ,"Patient"
Attribute VB_Ext_KEY = "Member15" ,"Patient"
Attribute VB_Ext_KEY = "Member16" ,"Rads"
Attribute VB_Ext_KEY = "Member17" ,"Validation"
Attribute VB_Ext_KEY = "Member18" ,"Validation"
Option Explicit

'change history
'scr 35348 - 30APR2003 - j.ray - removed check to not allow disposition for "PROC"

Public Enum EnumLookUps
    elEandM
    elCPT4
    elICD9
    elModifiersCPT4
    elModifiersEandM
    elAdminCode
    elDisposition
    elProviderIEN
    elProviderRole
    elApptType
    elMEPRS
End Enum

Public Enum EnumValidationMode
    vmNone              'skip all validation checks
    vmSaveToCollection  'save errors to a collection - for use in SADR
    vmImmediate         'immediately raise an error  - for use by Medcin Lite
End Enum

Public Enum EnumClassType
    ctCPT4Modifiers
    ctEandMModifiers
    ctAdminCodes
    ctEandMs
    ctEandM
End Enum

Private mvarPatient As Patient
Private mvarCPT4s As CPT4s
Private mvarEandMs As EandMs
Private mvarICD9s As ICD9s
Private mvarDisposition As Disposition
Private mvarEncounter As Encounter
Private mvarAppointment As Appointment
Private mvarValidation As Validation
Private mvarDirty As Boolean
Private mszDirty As String      'Compared to determine dirtyness
Private mvarConsults As Consults
Private mvarLabs As Labs
Private mvarMeds As Meds
Private mvarRads As Rads
Private mvarProblems As Problems
Private mvarAnP As AnP
Private mvarEncPregnancy As EncPregnancy
Private mvarEncAccident As EncAccident

'local variable(s) to hold property value(s)
Private mvarVersion As String 'local copy
Public Property Let Version(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Version = 5
    mvarVersion = vData
End Property


Public Property Get Version() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    Version = mvarVersion
End Property


Public Property Get Rads() As Rads
    If mvarRads Is Nothing Then
        Set mvarRads = New Rads
    End If


    Set Rads = mvarRads
End Property


Public Property Set Rads(vData As Rads)
    Set mvarRads = vData
End Property



Public Property Get Meds() As Meds
    If mvarMeds Is Nothing Then
        Set mvarMeds = New Meds
    End If


    Set Meds = mvarMeds
End Property


Public Property Set Meds(vData As Meds)
    Set mvarMeds = vData
End Property



Public Property Get Labs() As Labs
    If mvarLabs Is Nothing Then
        Set mvarLabs = New Labs
    End If


    Set Labs = mvarLabs
End Property


Public Property Set Labs(vData As Labs)
    Set mvarLabs = vData
End Property



Public Property Get Consults() As Consults
    If mvarConsults Is Nothing Then
        Set mvarConsults = New Consults
    End If


    Set Consults = mvarConsults
End Property


Public Property Set Consults(vData As Consults)
    Set mvarConsults = vData
End Property

'SCR 35618  Phil Crowder 5/5/03
Public Property Get Problems() As Problems
  If mvarProblems Is Nothing Then
    Set mvarProblems = New Problems
  End If
  
  Set Problems = mvarProblems
End Property

'SCR 35618  Phil Crowder 5/5/03
Public Property Set Problems(ByRef NewValue As Problems)
  Set mvarProblems = NewValue
End Property

Public Function ForceValidate() As Long
Attribute ForceValidate.VB_Description = "Returns number of Errors in the Errors collection. Performs overall validation, in order to catch things like a missing E&M code, or unset field values."
'The business rule is in the technical note we added - in order to complete the encounter,
'it must have at least one E&M code, one ICD-9 diagnosis code, and a disposition. grb. 11/21/02
'''The following specific messages should be added as they apply:
'''
'''If there is no diagnosis code
'''“A diagnosis has not been entered.”
'''If there is no E&M code
'''“An E&M code has not been entered.”
'''If there is no disposition code
'''“A disposition has not been entered.”
'''
'''If a procedure/order has not been associated with a diagnosis
'''“One or more procedures/orders have not been associated with a diagnosis.”
'''
'''If an E&M code has not been associated with a diagnosis
'''“One or more E&M codes have not been associated with a diagnosis.”

    Dim sz As String
    Dim objWAM2 As WAM
    Dim bRet As Boolean
    Dim objICD As ICD9
    Dim col As New Collection
    Dim objEM As EandM
    Dim objEM2 As EandM
    Dim objCPT As CPT4
    
    Dim p1 As Long
    Dim p2 As Long
    
    sz = PersistMe(vmNone)
    
    'this will catch errors like a missing disposition
    'should maintain any current errors
    Set objWAM2 = New WAM
    Set objWAM2.Validator.Lookups = Me.Validator.Lookups
    
    bRet = objWAM2.RestoreMe(sz, vmSaveToCollection)
    Set Validator.Errors = objWAM2.Validator.Errors
        
lblEMs:
    'this will catch E&M/ICD9 requirement
    If EandMs.Count = 0 Then
        Validator.Errors.Add vbObjectError + 100, "WAM.ForceValidate", "At least 1 E&M code must be present.", "An E&M code has not been entered."
    End If
    
    For p1 = 1 To EandMs.Count
        For p2 = 1 To EandMs.Count
            Set objEM = EandMs.Item(p1)
            Set objEM2 = EandMs.Item(p2)
            If p1 <> p2 Then
                If Len(objEM.EM_CODE) Then
                    If objEM.EM_CODE = objEM2.EM_CODE Then
                        Validator.Errors.Add vbObjectError + 113, "WAM.ForceValidate", "Duplicate E&M Codes are not allowed.", "Duplicate E&M Codes are not allowed."
                        GoTo ICDs
                    End If
                End If
            End If
        Next
    Next
        
ICDs:
    If ICD9s.Count = 0 Then
        Validator.Errors.Add vbObjectError + 101, "WAM.ForceValidate", "At least 1 ICD9 code must be present.", "At least 1 ICD9 code must be present."
    End If
    
    'No Dups on ICD Priorities
    On Error GoTo errNoICD9Dups
    Set col = New Collection
    For Each objICD In ICD9s
        If IsNumeric(objICD.ICD_DIAGNOSIS_PRIORITY) Then
            col.Add "X", CStr(objICD.ICD_DIAGNOSIS_PRIORITY)
        End If
    Next
            
Provider_Roles:
'Provider Roles
    With Appointment
        If Len(.Provider2_IEN) <> 0 Then
            If Len(.Provider2_Role_Code) = 0 Then
                Validator.Errors.Add vbObjectError + 110, "WAM.ForceValidate", "If Provider2_IEN is set, then Provider2_Role_Code must be set."
            End If
        End If
        If Len(.Provider3_IEN) <> 0 Then
            If Len(.Provider3_Role_Code) = 0 Then
                Validator.Errors.Add vbObjectError + 111, "WAM.ForceValidate", "If Provider3_IEN is set, then Provider3_Role_Code must be set."
            End If
        End If
    End With
    
'Disposition Code
    Select Case Me.Appointment.Appt_Type    'Limited to 5 Chars
        'scr 35348 - 30APR2003 - j.ray - removed check to not allow disposition for "PROC"
        'scr 35449 - 01MAY2003 - j.ray - added check to include "TCON"
        'probably should modify this to check against appt class = 3 instead
        Case "T-CON", "TCON*", "TCON$", "TCON"
            If Len(Me.Disposition.DISP_CODE) Then
                Validator.Errors.Add vbObjectError + 120, "WAM.ForceValidate", "For Telcon's the Disposition must be blank."
            End If
        Case Else
            If Me.Appointment.APPT_Class = "INPATIENT" Then
'           We can now enter inpatient diposistion codes SCR-35098
'                If Len(Me.Disposition.DISP_CODE) <> 0 Then
'                    Validator.Errors.Add vbObjectError + 121, "WAM.ForceValidate", "A disposition cannot be entered for an Inpatient Appointment.", "A disposition cannot be entered for an Inpatient Appointment."
'                End If
            Else
                If Len(Me.Disposition.DISP_CODE) = 0 Then
                    Validator.Errors.Add vbObjectError + 122, "WAM.ForceValidate", "A disposition has not been entered.", "A disposition has not been entered."
                End If
            End If
    End Select
'Priorities
    For Each objEM In Me.EandMs
        If objEM.EM_DIAGNOSIS_PRIORITY = vbNullString Then
            Validator.Errors.Add vbObjectError + 130, "WAM.ForceValidate", _
                "EM Code '" & objEM.EM_CODE & "' is not associated with a priority.", _
                "EM Code '" & objEM.EM_CODE & "' is not associated with a priority."
        End If
        If InStr(objEM.EM_DIAGNOSIS_PRIORITY, "0") Or InStr(objEM.EM_DIAGNOSIS_PRIORITY, "U") Then
            Validator.Errors.Add vbObjectError + 131, "WAM.ForceValidate", _
                "EM Code '" & objEM.EM_CODE & "' can only be associated with ICD's having priority 1-4.", _
                "EM Code '" & objEM.EM_CODE & "' can only be associated with ICD's having priority 1-4."
        End If
    Next
    For Each objCPT In Me.CPT4s
        If objCPT.CPT4_DIAGNOSIS_PRIORITY = vbNullString Then
            Validator.Errors.Add vbObjectError + 132, "WAM.ForceValidate", _
                "CPT Code '" & objCPT.CPT4_CODE & "' is not associated with a priority.", _
                "CPT Code '" & objCPT.CPT4_CODE & "' is not associated with a priority."
        End If
        If InStr(objCPT.CPT4_DIAGNOSIS_PRIORITY, "0") Or InStr(objCPT.CPT4_DIAGNOSIS_PRIORITY, "U") Then
            Validator.Errors.Add vbObjectError + 133, "WAM.ForceValidate", _
                "CPT Code '" & objCPT.CPT4_CODE & "' can only be associated with ICD's having priority 1-4.", _
                "CPT Code '" & objCPT.CPT4_CODE & "' can only be associated with ICD's having priority 1-4."
        End If
    Next
            
'ApptIEN
    If Len(Me.Appointment.APPT_IEN) = 0 Then
        Me.Validator.Errors.Add vbObjectError + 134, "MT_SADR.WritebackToADM", "Appt_IEN is missing for this Encounter."
    End If

XIT:
    ForceValidate = Validator.Errors.Count
    Exit Function
    
errNoICD9Dups:
    Validator.Errors.Add vbObjectError + 102, "WAM.ForceValidate", "Duplicate ICD_DIAGNOSIS_PRIORITYs are not allowed."
    Resume Provider_Roles
    
End Function

Public Function Dirty() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dirty
    If mszDirty <> PersistMe(vmNone) Then
        Dirty = True
    End If
    
End Function

Public Property Get Validator() As Validation
    Set Validator = mvarValidation
End Property

Public Property Get AnP() As AnP
  Set AnP = mvarAnP
End Property

Public Property Let AnP(vData As AnP)
  Set mvarAnP = vData
End Property

Public Property Get Appointment() As Appointment
    Set Appointment = mvarAppointment
End Property
Public Property Set Appointment(vData As Appointment)
    Set mvarAppointment = vData
End Property

Public Property Get Encounter() As Encounter
    Set Encounter = mvarEncounter
End Property
Public Property Set Encounter(vData As Encounter)
    Set mvarEncounter = vData
End Property

Public Property Get EncAccident() As EncAccident
    Set EncAccident = mvarEncAccident
End Property
Public Property Set EncAccident(vData As EncAccident)
    Set mvarEncAccident = vData
End Property

Public Property Get EncPregnancy() As EncPregnancy
    Set EncPregnancy = mvarEncPregnancy
End Property
Public Property Set EncPregnancy(vData As EncPregnancy)
    Set mvarEncPregnancy = vData
End Property

Public Property Get Disposition() As Disposition
    Set Disposition = mvarDisposition
End Property
Public Property Set Disposition(vData As Disposition)
    Set mvarDisposition = vData
End Property

Public Property Get ICD9s() As ICD9s
    Set ICD9s = mvarICD9s
End Property
Public Property Set ICD9s(vData As ICD9s)
    Set mvarICD9s = vData
End Property

Public Property Get CPT4s() As CPT4s

    Set CPT4s = mvarCPT4s
End Property
Public Property Set CPT4s(vData As CPT4s)
    Set mvarCPT4s = vData
End Property

Public Property Get EandMs() As EandMs

    Set EandMs = mvarEandMs
End Property
Public Property Set EandMs(vData As EandMs)
    Set mvarEandMs = vData
End Property

Public Property Set Patient(ByVal vData As Patient)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Patient = Form1
    Set mvarPatient = vData
End Property
Public Property Get Patient() As Patient
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Patient
    Set Patient = mvarPatient
End Property



''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mvarValidation = New Validation
    
    Set mvarAppointment = New Appointment
    mvarAppointment.Init mvarValidation
    
    Set mvarEncounter = New Encounter
    mvarEncounter.Init mvarValidation
    
    Set mvarEncAccident = New EncAccident
    mvarEncAccident.Init mvarValidation
    
    Set mvarEncPregnancy = New EncPregnancy
    mvarEncPregnancy.Init mvarValidation
    
    Set mvarDisposition = New Disposition
    mvarDisposition.Init mvarValidation
    
    Set mvarICD9s = New ICD9s
    mvarICD9s.Init mvarValidation
    
    Set mvarCPT4s = New CPT4s
    mvarCPT4s.Init mvarValidation
   
    Set mvarEandMs = New EandMs
    mvarEandMs.Init mvarValidation
    
    Set mvarPatient = New Patient
    mvarPatient.Init mvarValidation
        
    Set mvarProblems = New Problems
    mvarProblems.Init mvarValidation
    
    Set mvarRads = New Rads
    mvarRads.Init mvarValidation
    
    Set mvarMeds = New Meds
    mvarMeds.Init mvarValidation
    
    Set mvarLabs = New Labs
    mvarLabs.Init mvarValidation
    
    Set mvarConsults = New Consults
    mvarConsults.Init mvarValidation

    Set mvarAnP = New AnP
    
    mvarValidation.Init mvarICD9s, mvarCPT4s, mvarEandMs
    
End Sub
Private Sub Class_Terminate()
    
    mvarValidation.Release          'Possible Circular Reference
    Set mvarValidation = Nothing
    
    Set mvarAnP = Nothing
    Set mvarAppointment = Nothing
    Set mvarEncounter = Nothing
    Set mvarDisposition = Nothing
    Set mvarICD9s = Nothing
    Set mvarCPT4s = Nothing
    Set mvarEandMs = Nothing
    Set mvarPatient = Nothing
    Set mvarValidation = Nothing
    
    Set mvarProblems = Nothing
    Set mvarRads = Nothing
    Set mvarMeds = Nothing
    Set mvarLabs = Nothing
    Set mvarConsults = Nothing
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''
' Persistence Block
'''''''''''''''''''''''''''''''''''''''''''''
Public Function PersistMe(Optional ByVal ValidationMode As EnumValidationMode = vmNone) As String
    Dim lMode As EnumValidationMode
    
    lMode = Validator.ValidationMode
    Validator.ValidationMode = ValidationMode   'Disable Validation routines, '{{BOOKMARK Set to not None, to find missing validation routines}}
                                                'no need to persit descriptions
                                                'lookup tables may not be available
    PersistMe = Persist(Me)
    
    Validator.ValidationMode = lMode

End Function
Public Function RestoreMe(ByVal Persist As String, Optional ByVal ValidationMode As EnumValidationMode = vmNone) As Boolean
    Dim lMode As EnumValidationMode
    
    lMode = Validator.ValidationMode
    Validator.ValidationMode = ValidationMode   'Disable Validation routines,
                                                'lookup tables may not be available
    
    RestoreMe = Restore(Me, Persist)
    
    'Cache Current State For Future Dirty Checks
    mszDirty = Persist
    
    Validator.ValidationMode = lMode

End Function

Public Function ClassFactory(ByVal Flag As EnumClassType) As Object
    
    Dim objWam As WAM
    Dim objLookup As Lookup
    
    Set objWam = New WAM
    
    For Each objLookup In Me.Validator.Lookups
        objWam.Validator.Lookups.Add objLookup.Key, objLookup.Recordset
    Next
    
    Select Case Flag
    Case ctAdminCodes
        Set ClassFactory = objWam.Disposition.AdminCodes
        
    Case ctEandMs
        Set ClassFactory = objWam.EandMs
        
    Case ctEandM
        Set ClassFactory = objWam.EandMs.Add(1, 1, 1)
        
    Case ctEandMModifiers
         Set ClassFactory = objWam.EandMs.Add(1, 1, 1).Modifiers
   
    Case ctCPT4Modifiers
         Set ClassFactory = objWam.CPT4s.Add(1, 1, 1).Modifiers
         
    End Select
    
    objWam.Validator.ValidationMode = Me.Validator.ValidationMode
    
End Function
