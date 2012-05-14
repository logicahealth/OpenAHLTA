VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MMI_Sites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mCurr As MMI_Site
Private mvarIdx As Long                '- current addr in mcol

Public Function AddSite(ByRef oSite As MMI_Site) As Boolean
    On Error GoTo ErrHandler    '- just in case there's already one there
    If Not oSite Is Nothing Then
        mCol.Add oSite
        AddSite = True
        Set mCurr = oSite
        mvarIdx = mCol.Count
    End If
      
ErrHandler:
End Function

Public Function Add(ByVal sUpdateFlag As String, ByVal sFacilityNCID As String, _
            ByVal sMedRecNum As String, ByVal sRadiologyNum As String, _
            Optional ByVal sKey As String) As MMI_Site
    'create a new object
    Dim objNewMember As MMI_Site
    On Error GoTo ErrHandler
    
    Set objNewMember = New MMI_Site

    'set the properties passed into the method
    objNewMember.FacilityNCID = sFacilityNCID
    objNewMember.MedRecNum = sMedRecNum
    objNewMember.RadiologyNum = sRadiologyNum
    If LenB(sKey) > 0 Then
      mCol.Add objNewMember, sKey
    Else
      mCol.Add objNewMember
    End If

    'return the object created
    Set Add = objNewMember
    Set mCurr = objNewMember
    mvarIdx = mCol.Count
    Set objNewMember = Nothing
ErrHandler:

End Function

Public Property Get Item(ByRef vntIndexKey As Variant) As MMI_Site
  
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  Set mCurr = mCol(vntIndexKey)

End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(ByRef vntIndexKey As Variant)
    On Error Resume Next
    mCol.Remove vntIndexKey
    Set mCurr = New MMI_Site
    mvarIdx = 0
End Sub


Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mCurr = New MMI_Site
    mvarIdx = 0
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mCurr = Nothing
End Sub

Public Function MoveNext() As Boolean
    mvarIdx = mvarIdx + 1
    If mvarIdx <= mCol.Count Then
        Set mCurr = mCol.mvarIdx
        MoveNext = True
    Else
        Set mCurr = New MMI_Site
        mvarIdx = 0
        MoveNext = False
    End If
End Function

Public Function MoveFirst() As Boolean
    mvarIdx = 1
    If mvarIdx <= mCol.Count Then
        Set mCurr = mCol.mvarIdx
        MoveFirst = True
    Else
        Set mCurr = New MMI_Site
        mvarIdx = 0
        MoveFirst = False
    End If
 
End Function

