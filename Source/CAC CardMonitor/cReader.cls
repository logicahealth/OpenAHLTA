VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ReaderName As String
Public ReaderCurrentState As Long
Public ReaderEventState As Long
Public CardATR     As String

Public hCard      As Long
Public cardPresent As Boolean
Public CardContext As Long   '- constant for session
Public CardID     As String
Public CardState  As Long
Public OldCardID  As Long       '-- the card id that was here (just removed?)


Public FirstName         As String
Public LastName          As String
Public MiddleName        As String
Public DateOfBirth       As Date
Public Identifier        As String
Public IdentifierType    As Integer
Public ExpirationDate    As Date
Public Certificate       As String
Private m_PIN            As String

Public Event CardInserted(CUID As String)
Public Event CardRemoved(CUID As String)

Public Function StateChanged() As String
    '-- if inserted, return "INSERTED"
    '-- if removed , return "REMOVED
    If CardID = vbNullString And cardPresent Then
        '-- get the card
    End If
End Function

Public Property Let PIN(ByVal rhv As String)
    m_PIN = rhv
End Property


Friend Sub UpdateStatus(ByRef udtReaderCardStates As SCARD_READERSTATEA)


'-- Type SCARD_READERSTATEA
'    szReader        As String       ' reader name   (null-terminated)
'    pvUserData      As Long         ' user defined data (not used)
'    dwCurrentState  As Long         ' current state of reader at time of call
'    dwEventState    As Long         ' state of reader after state change
'    cbAtr           As Long         ' Number of bytes in the returned ATR
'    rgbAtr(35)      As Byte         ' Atr of inserted card, (extra alignment bytes)
'End Type
'Public Const SCARD_ABSENT               As Long = 1
'Public Const SCARD_PRESENT              As Long = 2

On Error GoTo errhandler
    Dim bcardPresent As Boolean
    bcardPresent = (udtReaderCardStates.dwEventState And SCARD_STATE_PRESENT) = SCARD_STATE_PRESENT
    ReaderCurrentState = udtReaderCardStates.dwEventState
    
    If bcardPresent And Not cardPresent Then
        'someone added a card
        cardPresent = True
        Call GetCardID(ReaderName, CardID)
        
        RaiseEvent CardInserted(CardID)
    ElseIf Not bcardPresent And cardPresent Then
        RaiseEvent CardRemoved(CardID)
        cardPresent = False
        CardID = ""
    End If
    Exit Sub

errhandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Friend Sub ResetPersonalInfo()
    On Error Resume Next

    FirstName = ""
    LastName = ""
    MiddleName = ""
    DateOfBirth = 0
    Identifier = ""
    IdentifierType = 0
    ExpirationDate = 0
    Certificate = ""
End Sub


