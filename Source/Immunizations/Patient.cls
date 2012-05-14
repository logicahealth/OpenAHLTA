VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFirstName As String 'local copy
Private mvarLastName As String 'local copy
Private mvarFMP As String 'local copy
Private mvarSSN As String 'local copy
Private mvarSponSSN As String 'local copy
Private mvarDOB As Date 'local copy
Private mvarIEN As Long 'local copy

'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarsUIC As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsService As String 'local copy
Private mvarsStatus As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsSex As String 'local copy
Private mvarsRank As String 'local copy
Private mvarsProvider As String
Private mvarsProviderNCID As String
Public Property Let sProvider(ByVal sProviderIn As String)
  On Error Resume Next
  mvarsProvider = sProviderIn
End Property

Public Property Get sProvider() As String
  On Error Resume Next
  sProvider = mvarsProvider
End Property
Public Property Let sProviderNCID(ByVal sProviderNCID_In As String)
  On Error Resume Next
  mvarsProviderNCID = sProviderNCID_In
End Property
Public Property Get sProviderNCID() As String
  On Error Resume Next
  sProviderNCID = mvarsProviderNCID
End Property

Public Property Let sRank(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.rank = 5
    mvarsRank = vData
End Property


Public Property Get sRank() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.rank
    sRank = mvarsRank
End Property



Public Property Let sSex(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sex = 5
    mvarsSex = vData
End Property


Public Property Get sSex() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sex
    sSex = mvarsSex
End Property



Public Property Let sStatus(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sStatus = 5
    mvarsStatus = vData
End Property


Public Property Get sStatus() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sStatus
    sStatus = mvarsStatus
End Property



Public Property Let sService(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sService = 5
    mvarsService = vData
End Property


Public Property Get sService() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sService
    sService = mvarsService
End Property



Public Property Let sUIC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUIC = 5
    mvarsUIC = vData
End Property


Public Property Get sUIC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUIC
    sUIC = mvarsUIC
End Property




Public Property Let lIEN(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IEN = 5
    mvarIEN = vData
End Property


Public Property Get lIEN() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IEN
   lIEN = mvarIEN
End Property



Public Property Let dDOB(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DOB = 5
    mvarDOB = vData
End Property


Public Property Get dDOB() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DOB
    dDOB = mvarDOB
End Property



Public Property Let sSponSSN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SponSSN = 5
    mvarSponSSN = vData
End Property


Public Property Get sSponSSN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SponSSN
    sSponSSN = mvarSponSSN
End Property



Public Property Let sSSN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SSN = 5
    mvarSSN = vData
End Property


Public Property Get sSSN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SSN
    sSSN = mvarSSN
End Property



Public Property Let sFMP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FMP = 5
    mvarFMP = vData
End Property


Public Property Get sFMP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FMP
    sFMP = mvarFMP
End Property



Public Property Let sLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastName = 5
    mvarLastName = vData
End Property


Public Property Get sLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastName
    sLastName = mvarLastName
End Property



Public Property Let sFirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarFirstName = vData
End Property


Public Property Get sFirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    sFirstName = mvarFirstName
End Property



