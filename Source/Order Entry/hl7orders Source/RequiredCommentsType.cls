VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RequiredCommentsType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarPrompt As String 'local copy
Private mvarRequired As Boolean 'local copy

Private mvarDefault_Value As String 'local copy
Private mvarHelp_Text As String 'local copy
Private mvarAnswer As String 'local copy
'local variable(s) to hold property value(s)
Private mvarAnswerCollection As Collection 'local copy
Public Enum hlRequiredCommentsDataTypes
    hlSet_of_codes
    hlFree_text
    hldatetime
    hlDateonly
    hlNumeric
    hlPointerToFile
End Enum
Public Enum hlSourceOfFileData
    hlSourceCHCSI = 1
    hlSourceCHCSII = 2
End Enum
Private mvarData_Type As hlRequiredCommentsDataTypes 'local copy
Private mvarUpperLimit As Variant 'local copy
Private mvarLowerLimit As Variant 'local copy
Private mvarMaxLength As String 'local copy
Private mvarRequiredCommentID As String 'local copy
Private mvarFileNumber As String
Private mvarSourceOfFileData As hlSourceOfFileData
Public Property Let RequiredCommentID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RequiredCommentID = 5
    mvarRequiredCommentID = vData
End Property


Public Property Get RequiredCommentID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RequiredCommentID
    RequiredCommentID = mvarRequiredCommentID
End Property
Public Property Let FileNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RequiredCommentID = 5
    mvarFileNumber = vData
End Property


Public Property Get FileNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RequiredCommentID
    FileNumber = mvarFileNumber
End Property
Public Property Let SourceOfFileData(ByVal vData As hlSourceOfFileData)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RequiredCommentID = 5
    mvarSourceOfFileData = vData
End Property


Public Property Get SourceOfFileData() As hlSourceOfFileData
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RequiredCommentID
    SourceOfFileData = mvarSourceOfFileData
End Property



Public Property Let MaxLength(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxLength = 5
    mvarMaxLength = vData
End Property


Public Property Get MaxLength() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxLength
    MaxLength = mvarMaxLength
End Property



Public Property Let LowerLimit(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LowerLimit = 5
    mvarLowerLimit = vData
End Property


Public Property Set LowerLimit(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LowerLimit = Form1
    Set mvarLowerLimit = vData
End Property


Public Property Get LowerLimit() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LowerLimit
    If IsObject(mvarLowerLimit) Then
        Set LowerLimit = mvarLowerLimit
    Else
        LowerLimit = mvarLowerLimit
    End If
End Property



Public Property Let UpperLimit(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpperLimit = 5
    mvarUpperLimit = vData
End Property


Public Property Set UpperLimit(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.UpperLimit = Form1
    Set mvarUpperLimit = vData
End Property


Public Property Get UpperLimit() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpperLimit
    If IsObject(mvarUpperLimit) Then
        Set UpperLimit = mvarUpperLimit
    Else
        UpperLimit = mvarUpperLimit
    End If
End Property






Public Property Set AnswerCollection(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AnswerCollection = Form1
    Set mvarAnswerCollection = vData
End Property


Public Property Get AnswerCollection() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AnswerCollection
    Set AnswerCollection = mvarAnswerCollection
End Property



Public Property Let ReturnedAnswer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Answer = 5
    mvarAnswer = vData
End Property


Public Property Get ReturnedAnswer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Answer
    ReturnedAnswer = mvarAnswer
End Property



Public Property Let Help_Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Help_Text = 5
    mvarHelp_Text = vData
End Property


Public Property Get Help_Text() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Help_Text
    Help_Text = mvarHelp_Text
End Property



Public Property Let Default_Value(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Default_Value = 5
    mvarDefault_Value = vData
End Property


Public Property Get Default_Value() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Default_Value
    Default_Value = mvarDefault_Value
End Property



Public Property Let Data_Type(ByVal vData As hlRequiredCommentsDataTypes)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Date_type = Form1
    mvarData_Type = vData
End Property


Public Property Get Data_Type() As hlRequiredCommentsDataTypes
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Date_type
    Data_Type = mvarData_Type
End Property



Public Property Let Required(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Required = 5
    mvarRequired = vData
End Property


Public Property Get Required() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Required
    Required = mvarRequired
End Property



Public Property Let Prompt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Prompt = 5
    mvarPrompt = vData
End Property


Public Property Get Prompt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Prompt
    Prompt = mvarPrompt
End Property



Private Sub Class_Initialize()
    Set mvarAnswerCollection = New Collection
    
End Sub
