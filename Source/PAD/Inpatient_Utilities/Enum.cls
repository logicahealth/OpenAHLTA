VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enumItemCritical
    isNone = 0
    isNormal = 1
    isAbnormal = 2
    isCritial = 3
End Enum

Public Enum enumItemStatus
    isPending = 0
    isPreliminary = 1
    isCompleted_Unaddr = 2
    isCompleted_Addr = 3
    isActive = 4
    isHold = 5
    isDiscontinued = 6
    isDispensed = 7
    isAll = 8
    isCompleted_Order = 9
End Enum

Public Enum enumItemAddMode
    InPatientOrder = 0
    SummaryOrder = 1
End Enum

Public Enum enumPriority
    PRIORITY_ROUTINE = 1
    PRIORITY_STAT = 2
    PRIORITY_ASAP = 3
    PRIORITY_NOTIFY = 4
    PRIORITY_PREOP = 5
End Enum


Public Enum enumPatientStatus
    isBoth = 0
    isInpatient = 1
    isOutpatient = 2
End Enum

Public Function GetStatus(ByVal eStatus As Integer) As String
    Select Case eStatus
        Case isPending
            GetStatus = "Pending"
        Case isPreliminary
            GetStatus = "Preliminary"
        Case isCompleted_Unaddr
            GetStatus = "Completed Unaddressed"
        Case isCompleted_Addr
            GetStatus = "Completed Addressed"
        Case isActive
            GetStatus = "Active"
        Case isHold
            GetStatus = "Hold"
        Case isDiscontinued
            GetStatus = "Discontinued"
        Case isDispensed
            GetStatus = "Dispensed"
        Case isAll
            GetStatus = "All"
        Case isCompleted_Order
            GetStatus = "Completed"
    End Select
End Function

Public Function GetPriority(ePriority As enumPriority) As String
     Select Case ePriority
          Case PRIORITY_ROUTINE: GetPriority = "Routine"
          Case PRIORITY_STAT: GetPriority = "STAT"
          Case PRIORITY_ASAP: GetPriority = "ASAP"
          Case PRIORITY_NOTIFY: GetPriority = "Notify and Routine"
          Case PRIORITY_PREOP: GetPriority = "Preop"
     End Select
End Function

Public Function GetNormalSatus(Normal As enumItemCritical) As String
     Select Case Normal
          Case isNone:     GetNormalSatus = "None"
          Case isNormal:   GetNormalSatus = "Normal"
          Case isAbnormal: GetNormalSatus = "Abnormal"
          Case isCritial:  GetNormalSatus = "Critical"
     End Select
End Function
