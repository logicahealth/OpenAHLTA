VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"GEMS_Address"
Attribute VB_Ext_KEY = "Member0" ,"GEMS_Address"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mCurr As GEMS_Address
Private mvarIdx As Long                '- current addr in mcol
Private mvarEOF As Boolean

Public Function AddAddress(oAddr As GEMS_Address) As Boolean
On Error GoTo errHandler    '- just in case there's already one there
  If Not oAddr Is Nothing Then
    If Not oAddr.AddrType = "" Then
      mCol.Add oAddr, CStr(oAddr.AddrType)
    Else
      mCol.Add oAddr
    End If
    AddAddress = True
    Set mCurr = oAddr
    mvarEOF = False
    mvarIdx = mCol.Count
  End If
  
errHandler:
End Function

Public Function Add(UpdateFlag As String, Country As String, County As String, Zip As String, State As String, City As String, Street2 As String, Street1 As String, AddrType As String) As GEMS_Address
    'create a new object
    Dim objNewMember As GEMS_Address
    Set objNewMember = New GEMS_Address
On Error GoTo errHandler

    'set the properties passed into the method
    objNewMember.UpdateFlag = UpdateFlag
    objNewMember.Country = Country
    objNewMember.County = County
    objNewMember.Zip = Zip
    objNewMember.State = State
    objNewMember.City = City
    objNewMember.Street2 = Street2
    objNewMember.Street1 = Street1
    objNewMember.AddrType = AddrType
    If Not AddrType = 0 Then
      mCol.Add objNewMember, CStr(AddrType)
    Else
      mCol.Add objNewMember
    End If

    'return the object created
    Set Add = objNewMember
    Set mCurr = objNewMember
    mvarEOF = False
    mvarIdx = mCol.Count
    Set objNewMember = Nothing
errHandler:

End Function

Public Property Get Item(vntIndexKey As Variant) As GEMS_Address
Attribute Item.VB_UserMemId = 0
  
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  Set mCurr = mCol(vntIndexKey)

End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
     On Error Resume Next
 mCol.Remove vntIndexKey
    Set mCurr = New GEMS_Address
    mvarIdx = 0
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mCurr = New GEMS_Address
    mvarIdx = 0
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mCurr = Nothing
End Sub

Public Function MoveNext() As Boolean
  mvarIdx = mvarIdx + 1
  If mvarIdx <= mCol.Count Then
    Set mCurr = mCol(mvarIdx)
    mvarEOF = False
    MoveNext = True
  Else
    Set mCurr = New GEMS_Address
    mvarEOF = True
    mvarIdx = 0
    MoveNext = False
  End If
End Function

Public Function MoveFirst() As Boolean
   mvarIdx = 1
  If mvarIdx <= mCol.Count Then
    Set mCurr = mCol.Item(mvarIdx)
    mvarEOF = False
    MoveFirst = True
  Else
    Set mCurr = New GEMS_Address
    mvarEOF = True
    mvarIdx = 0
    MoveFirst = False
  End If
 
End Function
'=============================== mcurr stuff ====================
Public Property Get UpdateFlag() As String:           UpdateFlag = mCurr.UpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mCurr.UpdateFlag = RHV:         End Property

Public Property Get AddrType() As String:           AddrType = mCurr.AddrType:    End Property
Public Property Let AddrType(ByVal RHV As String):  mCurr.AddrType = RHV:         End Property

Public Property Get Street1() As String:            Street1 = mCurr.Street1:  End Property
Public Property Let Street1(ByVal RHV As String):   mCurr.Street1 = RHV:      End Property

Public Property Get Street2() As String:            Street2 = mCurr.Street2:  End Property
Public Property Let Street2(ByVal RHV As String):   mCurr.Street2 = RHV:      End Property

Public Property Get City() As String:               City = mCurr.City:        End Property
Public Property Let City(ByVal RHV As String):      mCurr.City = RHV:         End Property

Public Property Get State() As String:              State = mCurr.State:      End Property
Public Property Let State(ByVal RHV As String):     mCurr.State = RHV:        End Property

Public Property Get Zip() As String:                Zip = mCurr.Zip:          End Property
Public Property Let Zip(ByVal RHV As String):       mCurr.Zip = RHV:          End Property

Public Property Get County() As String:             County = mCurr.County:    End Property
Public Property Let County(ByVal RHV As String):    mCurr.County = RHV:       End Property

Public Property Get Country() As String:            Country = mCurr.Country:  End Property
Public Property Let Country(ByVal RHV As String):   mCurr.Country = RHV:      End Property


