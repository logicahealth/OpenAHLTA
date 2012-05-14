VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "medOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDrugName As String 'local copy
Private mSynonyms As String 'local copy
Private mDefSig As String 'local copy
Private mDefRefill As Long 'local copy
Private mDefRoute As String 'local copy
Private mDefFreq As String 'local copy
Private mDefDose As String 'local copy
Private mDefForm As String
Private mDefQty As Long 'local copy
Private mCurrentStock As Long 'local copy
Private mActive As Boolean 'local copy
Private mIen As Long 'local copy
'Private mNDC As Collection
Private mNDC As String  '** 12/21/04 Frank - Change from Collection to String
Private mNSN As String
Private mNoteToProvider As String

Private mDosageStrength As String   '** 2/7/05 Frank
Private mContentUnit As String      '** 2/7/05 Frank

Public Property Let Active(ByVal vData As Boolean)
    mActive = vData
End Property

Public Property Get Active() As Boolean
    Active = mActive
End Property

Public Property Let Synonyms(ByVal vData As String)
    mSynonyms = vData
End Property

Public Property Get Synonyms() As String
    Synonyms = mSynonyms
End Property

Public Property Let DrugName(ByVal vData As String)
    mDrugName = vData
End Property

Public Property Get DrugName() As String
    DrugName = mDrugName
End Property

Public Property Let DefSig(ByVal vData As String)
    mDefSig = vData
End Property

Public Property Get DefSig() As String
    DefSig = mDefSig
End Property

Public Property Let DefRefill(ByVal vData As Long)
    mDefRefill = vData
End Property

Public Property Get DefRefill() As Long
    DefRefill = mDefRefill
End Property
Public Property Let DefRoute(ByVal vData As String)
    mDefRoute = vData
End Property

Public Property Get DefRoute() As String
    DefRoute = mDefRoute
End Property

Public Property Let DefFreq(ByVal vData As String)
    mDefFreq = vData
End Property

Public Property Get DefFreq() As String
    DefFreq = mDefFreq
End Property

Public Property Let DosageStrength(ByVal vData As String)
    mDosageStrength = vData
End Property

Public Property Get DosageStrength() As String
    DosageStrength = mDosageStrength
End Property

Public Property Let ContentUnit(ByVal vData As String)
    mContentUnit = vData
End Property

Public Property Get ContentUnit() As String
    ContentUnit = mContentUnit
End Property

Public Property Let DefDose(ByVal vData As String)
    mDefDose = vData
End Property

Public Property Get DefDose() As String
    DefDose = mDefDose
End Property

Public Property Let DefForm(ByVal vData As String)
    mDefForm = vData
End Property

Public Property Get DefForm() As String
    DefForm = mDefForm
End Property

Public Property Let DefQty(ByVal vData As Long)
    mDefQty = vData
End Property

Public Property Get DefQty() As Long
    DefQty = mDefQty
End Property

Public Property Let CurrentStock(ByVal vData As Long)
    mCurrentStock = vData
End Property

Public Property Get CurrentStock() As Long
    CurrentStock = mCurrentStock
End Property

Public Property Let Ien(ByVal vData As Long)
    mIen = vData
End Property

Public Property Get Ien() As Long
    Ien = mIen
End Property

'Tracker #43058
'Public Property Set NDC(ByVal oData As Collection)
'    Set mNDC = oData
'End Property
'Public Property Get NDC() As Collection
'    Set NDC = mNDC
'End Property
Public Property Let NDC(ByVal vData As String)
    mNDC = vData
End Property
Public Property Get NDC() As String
    NDC = mNDC
End Property

Public Property Let NSN(ByVal vData As String)
    mNSN = vData
End Property
Public Property Get NSN() As String
    NSN = mNSN
End Property

Public Property Let NoteToProvider(ByVal sData As String)
    mNoteToProvider = sData
End Property
Public Property Get NoteToProvider() As String
    NoteToProvider = mNoteToProvider
End Property
'End #43058
