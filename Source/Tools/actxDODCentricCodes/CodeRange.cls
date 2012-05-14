VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msCodeRangeStart   As String
Private msCodeRangeEnd     As String
Private meCodeSet          As CodeSetTypes
Private meCodeDesgination  As CodeDesignationTypes

Friend Property Get Range_Start_Code() As String
   Range_Start_Code = msCodeRangeStart
End Property

Friend Property Let Range_Start_Code(ByVal sNewValue As String)
   msCodeRangeStart = sNewValue
End Property

Friend Property Get Range_End_Code() As String
    Range_End_Code = msCodeRangeEnd
End Property

Friend Property Let Range_End_Code(ByVal sNewValue As String)
    msCodeRangeEnd = sNewValue
End Property

Friend Property Let CodeSet(ByVal eNewValue As CodeSetTypes)
    meCodeSet = eNewValue
End Property

Friend Property Get CodeSet() As CodeSetTypes
    CodeSet = meCodeSet
End Property

Friend Property Let CodeDesignation(ByVal eNewValue As CodeDesignationTypes)
    meCodeDesgination = eNewValue
End Property

Friend Property Get CodeDesignation() As CodeDesignationTypes
    CodeDesignation = meCodeDesgination
End Property

'Determine if the code is in this range
Public Function CodeInRange(ByVal sCode As String)
   
   'Strip out any spaces
   sCode = Trim(sCode)
   
   If Right(Range_Start_Code, 1) = "*" Then
      'Simple Wildcard Range - Compare only the left most characters
      CodeInRange = CBool(Left(sCode, Len(Range_Start_Code) - 1) = Left(Range_Start_Code, Len(Range_Start_Code) - 1))
   ElseIf sCode >= Range_Start_Code And sCode <= Range_End_Code Then
      CodeInRange = True
   End If
   
End Function
