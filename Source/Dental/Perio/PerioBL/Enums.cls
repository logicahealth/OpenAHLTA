'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eConstant
    perUnknownValue = -5765
End Enum

Public Enum eMeasurementLocation
    perLocOne = 1
    perLocTwo = 2
    perLocThree = 3
    perLocWholeTooth
End Enum

Public Enum eFurcation
    perFurcNone
    perperFurcOne
    perFurcTwo
    perFurcThree
    perFurcFour
End Enum

Public Enum eRootType
    perOneRoot
    perTwoRoot
    perThreeRoot
End Enum

Public Enum eMobility
    perMobNone
    perMobOne
    perMobTwo
    perMobThree
End Enum

Public Enum eEvalType
    perInitial
    perReEval
    perPostSurgical
    perMaintenance
    perConsultation
End Enum

Public Enum ePrognosis
    perUnknown
    perExcellent
    perGood
    perFair
    perQuestionable
    perPoor
    perHopeless
End Enum

Public Enum eColPosition
    perLeftEnd
    perRightEnd
    perTooth1
    perTooth16
    perTooth17
    perTooth32
End Enum

Public Enum eViewType
    perFullMouth
    perMaxArch
    perManArch
    perURQ
    perULQ
    perLRQ
    perLLQ
    perTextEntry
    perPrintGraphic
End Enum

Public Enum eNumRoots
    perOneRoot
    perTwoRoot
    perThreeRoot
End Enum

Public Enum eGridType
    perBuccal
    perLingual
End Enum

Public Enum eMeasurementType
    perCA
    perGM
    perProbDepth
    perBleeding
    perFurcation
    perMobility
    perPrognosis
    perBoneLoss
    perSuppuration
    perPlaque
    perMGJ
    perMissing
    perPontic
    perUnerupted
    perToothName
End Enum

Public Enum eToothSideType
    perSideBuccal
    perSideLingual
    perSideWholeTooth
End Enum
