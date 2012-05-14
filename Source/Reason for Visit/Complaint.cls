VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Complaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"complaint"
Option Explicit

Private msEncounterNumber As String  'what encounter is this related to
Private msEncounterFacilityNCID As String    'what facility was this complaint logged at
Private msComplaintText As String        'this is the text representation of the complaint, not necessary but nice to have so you dont have to look it up.
Private msSnoID As String         'this is the id of the complaint
Private miVisitType As Integer       'This is simply the representation of whether this is a new complaint or a followup to a previous complaint.
Private msComment As String          'Comment is split into 2 parts, the list that the complaint came from and the comment for the complaint itself
Private msPrefixCode As String       'This is the prefix that represents something like "Family history of" which would be and "F"
Private msStatusNCID As String           'This is whether the object is Active or should be Deleted
Private msBIIndex As String         'This is the body image index so that the user can tell which item on the Body Image relates to this complaint
Private msDataId As String      'This comes from the lookup in Problems, this is not set by the user when they create a new complaint

Public Enum enVisitType
    enVisitTypeNew = 0
    enVisitTypeFollowup = 1
End Enum

Public Property Get EncounterNumber() As String
    EncounterNumber = msEncounterNumber
End Property

Public Property Let EncounterNumber(number As String)
    msEncounterNumber = number
End Property

Public Property Get EncounterFacilityNCID() As String
    EncounterFacilityNCID = msEncounterFacilityNCID
End Property

Public Property Let EncounterFacilityNCID(number As String)
    msEncounterFacilityNCID = number
End Property

Public Property Get ComplaintText() As String
    ComplaintText = msComplaintText
End Property

Public Property Let ComplaintText(complaintString As String)
    msComplaintText = complaintString
End Property

Public Property Get SnoID() As String
    SnoID = msSnoID
End Property

Public Property Let SnoID(Id As String)
    msSnoID = Id
End Property

Public Property Get VisitTypeText() As String
    Select Case miVisitType
        Case 0
            VisitTypeText = "New"
        Case 1
            VisitTypeText = "Follow-Up"
    End Select
    
End Property

Public Property Get visittype() As enVisitType
    visittype = miVisitType
End Property

Public Property Let visittype(vType As enVisitType)
    miVisitType = vType
End Property

Public Property Get Comment() As String
    Comment = msComment
End Property

Public Property Let Comment(commentString As String)
    msComment = commentString
End Property

Public Property Get PrefixCode() As String
    PrefixCode = msPrefixCode
End Property

Public Property Let PrefixCode(code As String)
    msPrefixCode = PrefixCode
End Property

Public Property Get StatusNCID() As String
    StatusNCID = msStatusNCID
End Property

Public Property Let StatusNCID(NCID As String)
    msStatusNCID = NCID
End Property

Public Property Get BIIndex() As String
    BIIndex = msBIIndex
End Property

Public Property Let BIIndex(Location As String)
    msBIIndex = Location
End Property

Public Property Get DataId() As String
    DataId = msDataId
End Property

Public Property Let DataId(Id As String)
    msDataId = Id
End Property

