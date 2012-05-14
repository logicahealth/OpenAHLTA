VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IntrfcCHCSII_Dental.IDENTAL_DRC
Private mPatientLevelDRC As Integer
Private mPatLevelLstChgDte As Date
Private mPatLevelLstExmDte As Date
Private mEncLvlDRCLstExamDte As Date
Private mEncLvlDRCLstChgDte As Date
Private mEncLvlDRC As Integer

'*****************************************************************************
'Sub:SetEncounterLevelDRCInfo
'Author:KDunne
'Date:9/9/2002
'Purpose: Called from DentalEncounter, sets the Encounter Level DRC info after
'         the Encounter has been loaded.
'*****************************************************************************
Public Sub SetEncounterLevelDRCInfo(ByVal intDRC As Integer, ByVal dteLstExmDte As Date, _
            ByVal dteLstChgDate As Date)
      mEncLvlDRC = intDRC
      mEncLvlDRCLstExamDte = dteLstExmDte
      mEncLvlDRCLstChgDte = dteLstChgDate
      
End Sub
'*****************************************************************************
'Sub:SetPatLevelDRCInfo
'Author:KDunne
'Date:9/9/2002
'Purpose: Called from DentalEncounter, sets the Patient Level DRC info after
'         the Encounter has been loaded.
'*****************************************************************************
Public Sub SetPatLevelDRCInfo(ByVal intDRC As Integer, ByVal dteLstExmDte As Date, _
            ByVal dteLstChgDate As Date)
      mPatientLevelDRC = intDRC
      mPatLevelLstExmDte = dteLstExmDte
      mPatLevelLstChgDte = dteLstChgDate
End Sub


Private Property Get IDENTAL_DRC_EncLvlDRC() As Integer
    IDENTAL_DRC_EncLvlDRC = mEncLvlDRC
End Property

Private Property Get IDENTAL_DRC_EncLvlLstChgDte() As Date
    IDENTAL_DRC_EncLvlLstChgDte = mEncLvlDRCLstChgDte
End Property

Private Property Get IDENTAL_DRC_EncLvlLstExmDte() As Date
    IDENTAL_DRC_EncLvlLstExmDte = mEncLvlDRCLstExamDte
End Property

Private Property Get IDENTAL_DRC_PatLvlDRC() As Integer
    IDENTAL_DRC_PatLvlDRC = mPatientLevelDRC
End Property


Private Property Get IDENTAL_DRC_PatLvlLstChgDate() As Date
    IDENTAL_DRC_PatLvlLstChgDate = mPatLevelLstChgDte
End Property

Private Property Get IDENTAL_DRC_PatLvlLstExmDate() As Date
    IDENTAL_DRC_PatLvlLstExmDate = mPatLevelLstExmDte
End Property

Private Sub IDENTAL_DRC_UpdateDRC()

End Sub
