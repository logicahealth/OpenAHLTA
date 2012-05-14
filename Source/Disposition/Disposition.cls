VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Disposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Disposition"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'  Class Name: Disposition
'
'  Purpose: This class represents the template for the Disposition Object, which is
'           associated with an encounter in the 3M HEMS system.  It stores the data
'           for the disposition.
'
'  Properties:  (These are misnamed.  Prop
'    mlngFacilityNum (read/write) - The Facility NCID of the encounter the patient is
'                                   associated with.
'    mlngEncounterNum (read/write) - The NCID of the encounter the patient is assoicated
'                                    with.
'    mlngUnitNum (read/write) - The Unit NCID of the encounter the patient is assoicated
'                                    with.
'    mlngDispEventID (read/write) - The discharge disposition value of the encounter
'                                   that represents the EventID for the Clinical Text
'                                   object that holds the disposition data.
'    mlngDispositionNum (read/write) - The disposition selected by the user on the
'                                      disposition screen.
'    mlngDispositionDescripNum (read/write) - The disposition description number that
'                                             is linked to one of the three choices in
'                                             disposition below.
'    mstrDisposition (read/write) - This value can be one of the following:
'                                       1)  The time duration assoicated with the
'                                           Sick at Home/Quarters selection.
'                                       2)  The referred to text associated with
'                                           the Immediate Referral selection.
'                                       3)  The comment text associated with the Left
'                                           Against Medical Adivce selection.
'    mstrAdminOptions (read/write) - The administrative option on the disposition screen.
'    mlngReasonOptions (read/write) - The reason for visit option on the disposition screen.
'    mstrEMCode (read/write) - The E & M Code selection on the disposition screen.
'    mstrSpecialWorkStatus (read/write) - The Special Work Status selection on the
'                                         frmWorkStatus form.
'    mstrTelephoneConsult (read/write) - This is a temporary value to determine if the
'                                        encounter is a telephone consult code and should
'                                        be removed later.
'
'  Methods:
'********************************************************************************************

Option Explicit

Private mbAmbulatoryAuthorized As Boolean       'local copy

'Private miEncStatus As Long                  'local copy

'Private mlngFacilityNum As Long                 'local copy
'Private mlngEncounterNum As Long                'local copy
'Private mlngUnitNum As Long                     'local copy

Private mlngDispEventID As Long                 'local copy
'Private mlngDispositionNum As Long              'local copy
Private mlngDispositionDescripNum As Long       'local copy
'Private mstrDisposition As String               'local copy

Private mstrApptStatus As String                'local copy
'Private mstrAdminOptions As String              'local copy
'Private mlngReasonOptions As Long               'local copy

Private mstrEMCode As String                    'local copy
Private mstrEMNCID As String                    'local copy
Private mstrEMCalculation As String             'local copy
Private mstrEMCalcNote As String             'local copy
Private mstrEMDescription As String             'local copy
Private mstrMEPRSCode As String                 'local copy
Private mstrDMISCode As String                  'local copy
Private mstrPatientCategory As String                  'local copy
Private mstrACV As String                  'local copy

'Private msFollowupTimeFrame As String
'Private msFollowupComments As String
'Private msDiscussedComments As String
'Private msItemsDiscussed As String

Private mstrSpecialWorkStatus As String         'local copy

Private mstrTelephoneConsult As String          'local copy
Private mcolMyClinics As New Collection

' Dental Variables
'Private mstrDentalDispComments As String
'Private mstrDentalDispStatus As String
'Private mstrDentalDispStatusDesc As String
'Private mstrDentalDispOptions As String

' For theater only
'Private msDBNICategory As String
'Private msDBNICause As String
Private moEncDoc                As CHCSIIEncounterOps.Encounter
'moved out of the frmDisposition 11/2002
Public ProfileComment           As String
Public ProfileDiagnosisCode     As String
Public ProfileDiagnosisDesc     As String
Public ProfileStartDate         As Date
Public ProfileEndDate           As Date
Friend Property Set EncounterRef(eObject As Encounter)

    Set moEncDoc = eObject

End Property
Public Sub BuildClinicList(Optional bRetrieveAllClinics As Boolean = False)
    
    Dim objRecordset As New Recordset
    Dim oClinic As ClinicDef
    Dim sStmt As String
    
    On Error GoTo ErrHandler
    
    Set mcolMyClinics = New Collection
    
    If bRetrieveAllClinics Then
       sStmt = "Select c.NCID, c.Name" _
            & " From Clinic c" _
            & " Where c.facility_ncid = " & cwobjLogonEnviron.FacilityNCID _
            & " AND c.inactive <> 1" _
            & " Order by c.Name"
    Else
       sStmt = "Select c.NCID, c.Name" _
            & " From Clinic c, User_Clinic uc, Meprs_Code m" _
            & " Where uc.User_NCID = " & cwobjLogonEnviron.UserNCID _
            & " and c.ncid = uc.Clinic_NCID and m.IEN = c.Meprs_IEN" _
            & " and c.facility_ncid = " & cwobjLogonEnviron.FacilityNCID _
            & " and c.facility_ncid = m.facility_ncid" _
            & " and c.inactive <> 1 " _
            & " Order by c.Name"
    End If
    
    Set objRecordset = gobjCHCSII_DAS.OpenRecordset(sStmt)
    If Err Then
       GoTo ErrHandler
    ElseIf objRecordset Is Nothing Then
       GoTo CleanUp
    ElseIf (objRecordset.BOF And objRecordset.EOF) Then
       GoTo CleanUp
    End If
    
    Do While Not objRecordset.EOF
        Set oClinic = New ClinicDef
        oClinic.Name = objRecordset("NAME")
        oClinic.NCID = objRecordset("NCID")
        mcolMyClinics.Add oClinic, oClinic.NCID
        Set oClinic = Nothing
        objRecordset.MoveNext
    Loop
    GoTo CleanUp

ErrHandler:
    'Resume Next
    Err.Raise vbObjectError + 1, "Disposition:Disposition.BuildClinicList()", Err.Description

CleanUp:
    
    Set objRecordset = Nothing
    Set oClinic = Nothing
    
End Sub


'---implements the standard let/set methods
Public Property Let FacilityNum(ByVal plngFacilityNum As Long)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    'oEncDoc.FacilityNCID = plngFacilityNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 208  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get FacilityNum() As Long
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    FacilityNum = moEncDoc.FacilityNCID 'mlngFacilityNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 137  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let EncounterNum(ByVal plngEncounterNum As Long)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
        ' we aren't going to allow this WR 11/2002
    'oEncDoc.EncounterID = plngEncounterNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 138  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get EncounterNum() As Long
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    EncounterNum = moEncDoc.EncounterID 'mlngEncounterNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 139  'Pass this error back to the calling subroutine.
    End With

End Property

Public Sub ResetDefaults()
    
    mbAmbulatoryAuthorized = False
    
'    miEncStatus = 0
    
'    mlngFacilityNum = 0
'    mlngEncounterNum = 0
'    mlngUnitNum = 0
    
    mlngDispEventID = 0
'    mlngDispositionNum = 0
    mlngDispositionDescripNum = 0
    
'    mstrDisposition = ""
    mstrApptStatus = ""
'    mstrAdminOptions = ""
'    mlngReasonOptions = 0
    
    mstrEMCode = ""
    mstrEMNCID = ""
    mstrEMCalculation = ""
    mstrEMDescription = ""
    mstrMEPRSCode = ""
    mstrDMISCode = ""
    mstrPatientCategory = ""
    mstrACV = ""
    mstrSpecialWorkStatus = ""
    mstrTelephoneConsult = ""
    
'    msFollowupTimeFrame = ""
'    msFollowupComments = ""
'    msDiscussedComments = ""
'    msItemsDiscussed = ""
    
'    msDBNICategory = ""
'    msDBNICause = ""
    
End Sub

'---implements the standard let/set methods
Public Property Let UnitNum(ByVal plngUnitNum As Long)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    'mlngUnitNum = plngUnitNum
        'Let's not allow that WR
    'UnitNum = cwobjPatient.UnitNumber
        

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 140  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get UnitNum() As Long
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    UnitNum = cwobjPatient.UnitNumber  'mlngUnitNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 141  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let DispEventID(ByVal plngDispEventID As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mlngDispEventID = plngDispEventID

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 142  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get DispEventID() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    DispEventID = mlngDispEventID

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 143  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let Disposition(ByVal pstrDisposition As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.DispositionText = pstrDisposition

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 144  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get Disposition() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    Disposition = moEncDoc.DispositionText 'mstrDisposition

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 145  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let DispositionNum(ByVal plngDispositionNum As Long)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.DispositionNCID = plngDispositionNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 146  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get DispositionNum() As Long
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    DispositionNum = val(moEncDoc.DispositionNCID) 'mlngDispositionNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 147  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let DispositionDescripNum(ByVal plngDispositionDescripNum As Long)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mlngDispositionDescripNum = plngDispositionDescripNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 148  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get DispositionDescripNum() As Long
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    DispositionDescripNum = mlngDispositionDescripNum

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 149  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let AdminOptions(ByVal pstrAdminOptions As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.DispAdminOption = pstrAdminOptions

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 150  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get AdminOptions() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    AdminOptions = moEncDoc.DispAdminOption 'mstrAdminOptions

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 151  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let ReasonOptions(ByVal plngReasonOptions As Long)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.EncounterType = plngReasonOptions

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 152  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get ReasonOptions() As Long
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    ReasonOptions = moEncDoc.EncounterType 'mlngReasonOptions

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 153  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let EMCode(ByVal pstrEMCode As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    'This will be zero on a new encounter.
    
    mstrEMCode = pstrEMCode

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 154  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get EMCode() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    EMCode = mstrEMCode

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 155  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let SpecialWorkStatus(ByVal pstrSpecialWorkStatus As String)

    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrSpecialWorkStatus = pstrSpecialWorkStatus

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 156  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get SpecialWorkStatus() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    SpecialWorkStatus = mstrSpecialWorkStatus

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 157  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard let/set methods
Public Property Let TelephoneConsult(ByVal pstrTelephoneConsult As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrTelephoneConsult = pstrTelephoneConsult

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 158  'Pass this error back to the calling subroutine.
    End With

End Property

'---implements the standard get methods
Public Property Get TelephoneConsult() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    TelephoneConsult = mstrTelephoneConsult
    
    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 159  'Pass this error back to the calling subroutine.
    End With

End Property


Public Property Get AmbulatoryAuthorization() As Boolean
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    AmbulatoryAuthorization = mbAmbulatoryAuthorized

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 160  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let AmbulatoryAuthorization(ByVal bAmbulatoryValue As Boolean)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mbAmbulatoryAuthorized = bAmbulatoryValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 162  'Pass this error back to the calling subroutine.
    End With

End Property


Public Property Get DMISCode() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    DMISCode = mstrDMISCode

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 163  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let DMISCode(ByVal sDMISValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrDMISCode = sDMISValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 164  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get MEPRSCode() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    MEPRSCode = mstrMEPRSCode

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 165  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let MEPRSCode(ByVal sMEPRSValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrMEPRSCode = sMEPRSValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 166  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get EMNCID() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    EMNCID = moEncDoc.EandMNCID 'mstrEMNCID

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 167  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let EMNCID(ByVal sEMNCIDValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.EandMNCID = sEMNCIDValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 168  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get ApptStatus() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    ApptStatus = mstrApptStatus

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 169  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let ApptStatus(ByVal sApptStatusValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrApptStatus = sApptStatusValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 170  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get EncStatus() As Integer
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    EncStatus = moEncDoc.Status 'miEncStatus

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 177  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let EncStatus(ByVal iEncStatusValue As Integer)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.Status = iEncStatusValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 178  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get EMDescription() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    EMDescription = mstrEMDescription

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 179  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let EMDescription(ByVal sEMDescValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrEMDescription = sEMDescValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 180  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get EMCalculation() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    EMCalculation = moEncDoc.EandMCalc 'mstrEMCalculation

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 181  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let EMCalculation(ByVal sEMCalcValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.EandMCalc = sEMCalcValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 182  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get PatientCategory() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    PatientCategory = mstrPatientCategory

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 183  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let PatientCategory(ByVal sCategoryValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrPatientCategory = sCategoryValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 184  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get ACV() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    ACV = mstrACV

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 185  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let ACV(ByVal sACVValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    mstrACV = sACVValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 186  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get FollowUpTimeFrame() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    FollowUpTimeFrame = moEncDoc.FollowUpTimeFrame 'msFollowupTimeFrame

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 187  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let FollowUpTimeFrame(ByVal sFollowUpValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.FollowUpTimeFrame = sFollowUpValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 188  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get FollowupComments() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    FollowupComments = moEncDoc.FollowupComments 'msFollowupComments

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 189  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let FollowupComments(ByVal sCommentsValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    'SCR-7464   08/27/2000  H Stinson  ADDED:
    If Len(sCommentsValue$) > FLD_LEN__FOLLOWUP_COMMENTS Then
    
        MsgBxARMd "Truncating ""followup comments"" to max field length: " _
             & FLD_LEN__FOLLOWUP_COMMENTS + "."
        sCommentsValue$ = Trim$(Left$(sCommentsValue$, FLD_LEN__FOLLOWUP_COMMENTS))
    
    End If
    moEncDoc.FollowupComments = sCommentsValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 190  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get ItemsDiscussed() As String
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    ItemsDiscussed = moEncDoc.ItemsDiscussed 'msItemsDiscussed

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 191  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let ItemsDiscussed(ByVal sItemsValue As String)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    moEncDoc.ItemsDiscussed = sItemsValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 192  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Get MyClinics() As Collection
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    Set MyClinics = mcolMyClinics

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 193  'Pass this error back to the calling subroutine.
    End With

End Property

Public Property Let MyClinics(ByVal cClinicValue As Collection)
    
    On Error GoTo ErrHandler    'NOTE: This subr. does not need an error handler,
                                'but one is added just to keep the automatic code
                                'reviewer tool from complaining.
    
    Set mcolMyClinics = cClinicValue

    Exit Property
    
ErrHandler:
    With Err
        .Description = .Description + vbLf _
                     + "Error setting property "    '+ <Add property name here>.
        .Raise 194  'Pass this error back to the calling subroutine.
    End With

End Property


Public Property Get EMCalcNote() As String
    EMCalcNote = moEncDoc.EMCalcOverrideNote 'mstrEMCalcNote
End Property

Public Property Let EMCalcNote(ByVal sNoteValue As String)
    moEncDoc.EMCalcOverrideNote = sNoteValue
End Property

Public Property Get DBNICategory() As String
    DBNICategory = moEncDoc.DBNICategory 'msDBNICategory
End Property

Public Property Let DBNICategory(ByVal sCategory As String)
    moEncDoc.DBNICategory = sCategory
End Property

Public Property Get DBNICause() As String
    DBNICause = moEncDoc.DBNICause 'msDBNICause
End Property

Public Property Let DBNICause(ByVal sCause As String)
    moEncDoc.DBNICause = sCause
End Property

Public Property Get DentalDispStatus() As String
    
    DentalDispStatus = moEncDoc.DentalDispStatus 'mstrDentalDispStatus
    
End Property

Public Property Let DentalDispStatus(ByVal sValue As String)

    moEncDoc.DentalDispStatus = sValue

End Property


Public Property Get DentalDispStatusDesc() As String

    DentalDispStatusDesc = moEncDoc.DentalDispStatusDesc 'mstrDentalDispStatusDesc

End Property

Public Property Let DentalDispStatusDesc(ByVal sValue As String)

    moEncDoc.DentalDispStatusDesc = sValue
    
End Property

Public Property Get DentalDispOptions() As String

    DentalDispOptions = moEncDoc.DentalDispOptions 'mstrDentalDispOptions
    
End Property

Public Property Let DentalDispOptions(ByVal sValue As String)

    moEncDoc.DentalDispOptions = sValue

End Property

Public Property Get DentalDispComments() As String

    DentalDispComments = moEncDoc.DentalDispComments 'mstrDentalDispComments

End Property

Public Property Let DentalDispComments(ByVal sValue As String)

    moEncDoc.DentalDispComments = sValue
    
End Property

Private Sub Class_Initialize()
    '--- SCR-ARM-16000 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForDisposition
End Sub

Public Property Get DiscussedComments() As String
    DiscussedComments = moEncDoc.DiscussedComments 'msDiscussedComments
End Property

Public Property Let DiscussedComments(ByVal RHV As String)
    
    If Len(RHV) > FLD_LEN__FOLLOWUP_COMMENTS Then
        MsgBxARMd "Truncating ""discussed comments"" to max field length: " & FLD_LEN__FOLLOWUP_COMMENTS & "."
        RHV = Trim$(Left$(RHV, FLD_LEN__FOLLOWUP_COMMENTS))
    End If
    moEncDoc.DiscussedComments = RHV

End Property
