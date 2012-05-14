VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedDispense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Display(OrderID As String, _
                    OrderIEN As String, _
                    oUser As ICHCSII_User, _
                    oComm As MMMHISComm.Comm, _
                    oLogon As MMMHISLogon.Logon, _
                    oConfig As MMMHISConfiguration.Config, _
                    oPatient As MMMHISPatient.Patient, _
                    bIsPending As Boolean)
    

    

    Dim bDispense As Boolean
    
    On Error GoTo Err_Handler
    
    If goBB Is Nothing Then Set goBB = New Blackboard
    
    Set goBB.User = oUser
    Set goBB.Comm = oComm
    Set goBB.Logon = oLogon
    Set goBB.Patient = oPatient
    Set goBB.Config = oConfig
    
    Load frmDispense

    
    Call frmDispense.Display(OrderID, bDispense, bIsPending, False, True)
    
    'SF SCR 76434
    Unload frmDispense
    
    Exit Sub
Err_Handler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsMedDispense.Display", App.Title, vbExclamation
    
End Sub



Public Function Display2(OrderID As String, _
                    bDispense As Boolean, _
                    RxSig As String, _
                    RxComments As String, _
                    Optional CallFromGemsOE As Boolean = False, _
                    Optional CallFromLabelPrinting As Boolean = False, _
                    Optional ByRef IsAutoPrintLabel As Boolean = False, _
                    Optional DefaultPrinter As String, _
                    Optional LabelPrinter As String, _
                    Optional ByRef IsAutoPrint As Boolean = False) ' Quan 78763
    
    Dim bIsPending As Boolean
    
    On Error GoTo Err_Handler

''    If Printer.DeviceName <> DefaultPrinter Then Call SetPrinterDefault(DefaultPrinter)
    
    bIsPending = bDispense
    
    Load frmDispense
    
    frmDispense.CallFromLabelPrinting = CallFromLabelPrinting
    frmDispense.LabelPrinter = LabelPrinter
    frmDispense.DefaultPrinter = DefaultPrinter
    
    'SF SCR 76434
    Display2 = frmDispense.Display(OrderID, bDispense, True, CallFromGemsOE, False, CallFromLabelPrinting, RxSig, RxComments)
    Unload frmDispense
    
    IsAutoPrintLabel = frmDispense.AutoPrintLabel
    IsAutoPrint = frmDispense.chkAutoPrint 'Quan 78763
    
    Exit Function
Err_Handler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsMedDispense.Display2", App.Title, vbExclamation
    
End Function
