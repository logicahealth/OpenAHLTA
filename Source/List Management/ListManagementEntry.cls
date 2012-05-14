VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarSnoId As String 'local copy
Private mvarPrefix As String 'local copy
Private mvarDescription As String 'local copy
Private mvarNCID As Long 'local copy
Private mvarCode As String 'local copy
Private mvarClinicNCID As Long 'local copy
Private mvarKey As String
Private mvarNotFromMedcin As String    'JT SCR 42399 - Added new property to indicate we don't have a medcin id and must use the code value for the description (diagnosis only)

Friend Sub Init(ByVal SnoID As String, ByVal PREFIX As String, ByVal DESCRIPTION As String, _
                ByVal NCID As Long, ByVal Code As String, _
                ByVal ClinicNCID As Long, ByVal Key As String, Optional ByVal NotFromMedcin As Boolean = False)

    mvarSnoId = SnoID
    mvarPrefix = PREFIX
    mvarDescription = DESCRIPTION
    mvarNCID = NCID
    mvarCode = Code
    mvarClinicNCID = ClinicNCID
    mvarKey = Key
    mvarNotFromMedcin = NotFromMedcin   'JT SCR 42399 - Added new property
    
End Sub
Public Property Get SnoID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnoID
    SnoID = mvarSnoId
End Property


Public Property Get PREFIX() As String
'used when retrieving value of a property, on the right side of an assignment.
    PREFIX = mvarPrefix
End Property

Public Property Get DESCRIPTION() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    DESCRIPTION = mvarDescription
End Property

Public Property Get ClinicNCID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClinicNCID
    ClinicNCID = mvarClinicNCID
End Property


Public Property Get Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property

Public Property Get NCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NCID
    NCID = mvarNCID
End Property

Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NCID
    Key = mvarKey
End Property

'JT BEGIN SCR 42399 - Added new property
Public Property Get NotFromMedcin() As Boolean
   NotFromMedcin = mvarNotFromMedcin
End Property
'JT END SCR 42399
