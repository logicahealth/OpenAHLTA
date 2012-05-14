VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public MnuButtonID As Integer 'The ID of the menuitem's parent (a menu)
Public MnuButtonName As String
Public MnuSection As Boolean 'Single Tooth (F) or Multi (T)
Public MnuGroup As eGroup
Public mnuItemId As Integer 'The key for this menuitem
Public MnuItemName As String
Public MnuItemCodeID As Integer 'The ADA CodeID
Public MnuItemDRC As String
Public MnuItemVisible As Boolean
Public MnuItemToInstructionID As Long 'The ID of the menuitem's instruction set
Public MnuItemInstructions As colMenuItemInstrs
Public SubMnuItemID As Long 'The ID of the menuitem's sub menu

'Public MenuId As Integer
Public BizRule As eBizRulz
'Public MenuItemID As Integer
Public BizSurface As Integer
Public BizSurfacePlus As Integer
Public BizToothSequence As eToothSequence
Public BizToothPosition As eToothPosition
Public BizToothType As eToothType
Public BizCodeID As Integer
Public BizToothEndo As eBizToothEndo
Public BizToothMax As eMaxOrMan
Public BizImplantOrTooth As eBizImplantOrTooth
Public BizToothRoots As eToothRoots
Public BizToothSupernumerary As eBizToothSupernumerary

Public Enum eBizToothSupernumerary
    eTrue = -1
    eFalse = 0
    eAny = 777
End Enum

Public Enum eBizImplantOrTooth
    eImplant = -1
    eTooth = 0
    eAny = 777
End Enum

Public Enum eBizToothEndo
    eTrue = -1
    eFalse = 0
    eAny = 777
End Enum

Public Enum eEndodonicTherapy
  eEndo = -1
  eRetreat = 0
  eAny = 777
End Enum

Public Enum eMaxOrMan
  eMaxillary = -1
  eMandibular = 0
  eAny = 777
End Enum

Public Enum eToothSequence
  ePrimary = -1
  eSecondary = 0
  eAny = 777
End Enum

Public Enum eToothPosition
  ePosterior = -1
  eAnterior = 0
  eAny = 777
End Enum

Public Enum eToothType
  ePreMolar = -1
  eMolar = 0
  'eBicuspid = 2
  eAny = 777
End Enum

Public Enum eToothRoots
    eMultiple = -1
    eOne = 0
    eAll = 777
End Enum

Public Enum eBizRulz
  eSurface = 0
  eSurfaceNoWorkLoad = 17
  eSurfaceToothSeq = 1
  eSurfaceToothSeqPos = 2
  eExtract = 3
  eToothSeqPos = 4
  eToothSeqOnly = 5
  eMissing = 6
  eToothPosOnly = 7
  eToothSeqType = 8
  eToothTypeOnly = 9
  eMultiMissing = 10
  eImplants = 11
  eMaxMan = 12
  eToothTypePos = 13
  eImplantOnly = 14
  eToothOnly = 19
  eGeneric = 15 'This replaces most of the biz rules
  eNone = 99
  eDisabled = 100

End Enum
Public Function GetMenuInstruction(StatusID As Long) As clsMenuItemInstr
    
'This function returns the correct menuinstruction based on the status and menuitemid
    
On Error GoTo ErrHnd

Dim vGroup As eGroup
    
    'Translate the StatusID into the group
    Select Case StatusID
        
        'All the Dx statuses
        Case DxStatusEnum.dxActiveFinding, DxStatusEnum.dxActiveNew, DxStatusEnum.dxActiveObserve, _
            DxStatusEnum.dxActiveOngoing, DxStatusEnum.dxActiveProvisional, DxStatusEnum.dxNonActiveConfirmedNegative, _
            DxStatusEnum.dxNonActiveDeleted, DxStatusEnum.dxNonActiveResolved, DxStatusEnum.dxNonActiveUnableToConfirm
                
            vGroup = eDx
            
        Case TxStatusEnum.txPlanned, TxStatusEnum.txInProgress
            vGroup = eTxP
        
        Case TxStatusEnum.txCompleted, TxStatusEnum.txPreExisting
            vGroup = eTx
        
        Case TxStatusEnum.txRevised
            'Use the default group
            vGroup = MnuGroup
    End Select
    
    'Get the menuinstruction
    Set GetMenuInstruction = goMenuInstructions(MnuItemToInstructionID & "|" & vGroup)
    
Exit Function

ErrHnd:
    MsgBxARMd Err.Description & " source:clsMenuItem.GetMenuInstruction. Could not find the menu instruction in the collection based on the MnuItemToInstructionID of: " & MnuItemToInstructionID & " and the status of: " & StatusID & ".", vbCritical
    Set GetMenuInstruction = Nothing
    
End Function
Private Sub Class_Initialize()

    Set MnuItemInstructions = New colMenuItemInstrs

End Sub

Private Sub Class_Terminate()
Set MnuItemInstructions = Nothing
End Sub
