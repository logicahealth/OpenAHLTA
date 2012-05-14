VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_Generic_Datas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"GEMS_Generic_Data"
Attribute VB_Ext_KEY = "Member0" ,"GEMS_Generic_Data"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mvarOrdinal As Long


Public Function AddGenericData(oGenD As GEMS_Generic_Data)
  mvarOrdinal = mvarOrdinal + 1
  '*-- Coleman 051801
  oGenD.Ordinal = mvarOrdinal
  Debug.Print "AddGenericData: " & oGenD.Ordinal & "  Type: " & oGenD.Type_NCID & "  Value_Text: " & oGenD.Value_Text & "   Value_NCID: " & oGenD.Value_NCID
  mCol.Add oGenD
End Function

'*-- Coleman 061201: Added sUpdateFlag parameter
Public Function Add(sUpdateFlag As String, Type_NCID As Long, Error_Text As String, Value_Text As String, Value_RSID As Long, Value_NCID As Long, Parent As Long, Unit_Number As Long) As GEMS_Generic_Data
    'create a new object
    Dim objNewMember As GEMS_Generic_Data
    Set objNewMember = New GEMS_Generic_Data
    On Error GoTo errHandler
    
    'set the properties passed into the method
    '*-- Coleman 061201
    mvarOrdinal = mvarOrdinal + 1
    objNewMember.Ordinal = mvarOrdinal
    
    objNewMember.UpdateFlag = sUpdateFlag   '*-- Coleman 061201
    objNewMember.Type_NCID = Type_NCID
    objNewMember.Error_Text = Error_Text
    objNewMember.Value_Text = Value_Text
    objNewMember.Value_RSID = Value_RSID
    objNewMember.Value_NCID = Value_NCID
    objNewMember.Parent = Parent
    objNewMember.Unit_Number = Unit_Number
        
    mCol.Add objNewMember
    Debug.Print "GenericDatas.Add: " & CStr(mvarOrdinal) & "  Type: " & Type_NCID & "  Value_Text: " & Value_Text & "   Value_NCID: " & Value_NCID
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
errHandler:

End Function

Public Property Get Item(Optional vntKey As Variant) As GEMS_Generic_Data
Attribute Item.VB_UserMemId = 0
  Static key As Variant
  Static LastCnt As Long
  Dim Cnt As Long
  Dim oGenD As GEMS_Generic_Data
  
  If Not IsEmpty(vntKey) Then
    key = vntKey
    LastCnt = 0
  Else
    If key = "" Then
      Exit Property
    End If
  End If
  
  Cnt = 0
  For Each oGenD In mCol
    '*-- Coleman 051801: If Cnt starts at 0 , the first object
    '*-- never gets processed.
    Cnt = Cnt + 1
    If oGenD.Type_NCID = key And Cnt > LastCnt Then
      Set Item = oGenD
      Set oGenD = Nothing
      LastCnt = Cnt
      Exit Property
    End If
    '*-- Coleman 051801
    'Cnt = Cnt + 1
  Next oGenD
    
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
  On Error Resume Next

    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


