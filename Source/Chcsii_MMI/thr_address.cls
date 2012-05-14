VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "thr_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IAddress

Private moAddress   As GEMS_Addresses

Private mvaridx     As Long
Private BOF         As Boolean
Private EOF         As Boolean

Friend Property Set GemsAddress(oAddress As GEMS_Addresses)

    Set moAddress = oAddress
    BOF = True
    mvaridx = 0
    
    If moAddress.Count = 0 Then
        EOF = True
    Else
        EOF = False
    End If
    
End Property

Private Property Get IAddress_BOF() As Boolean

    IAddress_BOF = BOF


End Property

Private Property Let IAddress_City(ByVal RHS As String)

    moAddress.City = RHS

End Property

Private Property Get IAddress_City() As String

    IAddress_City = moAddress.City

End Property

Private Property Let IAddress_Country(ByVal RHS As String)

    moAddress.Country = RHS

End Property

Private Property Get IAddress_Country() As String

    IAddress_Country = moAddress.Country

End Property

Private Property Let IAddress_County(ByVal RHS As String)

    moAddress.County = RHS

End Property

Private Property Get IAddress_County() As String

    IAddress_County = moAddress.County

End Property

Private Property Get IAddress_EOF() As Boolean

    IAddress_EOF = EOF

End Property

Private Sub IAddress_GotoFirst()

    mvaridx = 1
    If moAddress.Count = 0 Then
        EOF = True
        BOF = True
    ElseIf mvaridx <= moAddress.Count Then
        EOF = False
        BOF = False
        moAddress.MoveFirst
    End If

    'Debug.Assert mvaridx > 0

End Sub

Private Sub IAddress_GotoLast()
    Dim pos As Long
    
    mvaridx = moAddress.Count
    
    pos = mvaridx
    moAddress.MoveFirst
    Do While pos > 1
        moAddress.MoveNext
        pos = pos - 1
    Loop
    
    
End Sub

Private Sub IAddress_GotoNext()

    mvaridx = mvaridx + 1
    If mvaridx > moAddress.Count Then
        EOF = True
        Exit Sub
    End If
    
    Debug.Assert mvaridx <= moAddress.Count
    moAddress.MoveNext

End Sub

Private Sub IAddress_GotoPrevious()

End Sub

Private Property Let IAddress_PostCode(ByVal RHS As String)

    moAddress.Zip = RHS

End Property

Private Property Get IAddress_PostCode() As String

    IAddress_PostCode = moAddress.Zip

End Property

Private Property Let IAddress_State(ByVal RHS As String)

    moAddress.State = RHS

End Property

Private Property Get IAddress_State() As String

    IAddress_State = moAddress.State

End Property

Private Property Get IAddress_StoreTime() As Date

End Property

Private Property Let IAddress_StoreTime(ByVal RHS As Date)

End Property

Private Property Let IAddress_Street(ByVal RHS As String)

    moAddress.Street1 = RHS

End Property

Private Property Get IAddress_Street() As String

    IAddress_Street = moAddress.Street1

End Property

Private Property Let IAddress_Street2(ByVal RHS As String)

    moAddress.Street2 = RHS

End Property

Private Property Get IAddress_Street2() As String

    IAddress_Street2 = moAddress.Street2

End Property

Private Property Let IAddress_Type(ByVal RHS As String)
    Dim gemsAddr As GEMS_Address
    Dim v       As Variant

    If mvaridx = 0 Then
        IAddress_GotoLast
    End If
    
    If mvaridx > moAddress.Count Or moAddress.Count = 0 Then
        Set gemsAddr = New GEMS_Address
     
        gemsAddr.UpdateFlag = "N"
        gemsAddr.AddrType = RHS
     
        moAddress.AddAddress gemsAddr
        mvaridx = moAddress.Count
        IAddress_GotoLast
    End If

    moAddress.AddrType = RHS
    'Type 0 means delete, 3M documentation states
    If Val(RHS) = 0 Then
        moAddress.Item(mvaridx).UpdateFlag = "D"
    End If
    

End Property

Private Property Get IAddress_Type() As String

    IAddress_Type = moAddress.AddrType

End Property
