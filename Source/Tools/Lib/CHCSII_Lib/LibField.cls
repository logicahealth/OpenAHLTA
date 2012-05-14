VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        clsLibField/LibField.cls - Data field object
'''
'''Description:   Allows to define, store, expose to user and evaluate single data element
'''
'''Modules Required:
'''
'''Note:
'''
'''Warning:          First assignment of the value is interpreted as indication that
'''                  object is ready to be used therefor all of the data definition
'''                  GUI properties should be set prior to setting the value
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME           As String = "mclsLibField."

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mbInit                As Boolean

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mobjItem              As mclsLibItem
Attribute mobjItem.VB_VarHelpID = -1
Private WithEvents mobjGUI    As mclsLibFieldGUI
Attribute mobjGUI.VB_VarHelpID = -1
Private mobjData              As mclsLibFieldData

Private Function M_bSetDefaultFormat() As Boolean
'''************************************************************************
'''Called By:     P_bInit
'''
'''Arguments:
'''
'''Description:   Based on the current data type set default output format
'''               according to DOD style guide
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - True
'''               Else           - False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bSetDefaultFormat"

   Dim dtpData          As DTPicker
   
On Error GoTo ErrHnd

      '''If default format is requested
10    If Len(mobjGUI.psFormat) = 0 Then
20       Select Case mobjData.plType
            Case lDTYP_Integer, lDTYP_Long
40             mobjGUI.psFormat = ""
            Case lDTYP_Double
60             mobjGUI.psFormat = "0.00"
            Case lDTYP_Date
80             mobjGUI.psFormat = gobjLibFormat.MskDspDate
            Case lDTYP_DateTime
100            mobjGUI.psFormat = gobjLibFormat.MskDspDateTime
            Case lDTYP_Time
120            mobjGUI.psFormat = gobjLibFormat.MskDspTime
         End Select
400      If Not mobjGUI.pctlData Is Nothing Then
500         If TypeOf mobjGUI.pctlData Is DTPicker Then
520            Set dtpData = mobjGUI.pctlData
540            dtpData.Format = dtpCustom
600            Select Case mobjData.plType
                  Case lDTYP_Date
620                  dtpData.CustomFormat = "dd MMM yyyy"
                  Case lDTYP_DateTime
640                  dtpData.CustomFormat = "dd MMM yyyy hh:mm:ss"
                  Case lDTYP_Time
660                  dtpData.UpDown = True
680                  dtpData.CustomFormat = "hh:mm:ss"
               End Select
               DoEvents
            End If
         End If
      End If

999   M_bSetDefaultFormat = True
   
Exit Function
ErrHnd:
   M_bSetDefaultFormat = False
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function


Private Function M_bSetMinMax() As Boolean
'''************************************************************************
'''Called By:     P_bInit
'''
'''Arguments:
'''
'''Description:   Based on the current data type set maximum input length
'''
'''Note:          this procedure is called only if current data control is text box
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - True
'''               Else           - False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bSetMinMax"

   Dim lMaxLen          As Long
   Dim lLenMinVal       As Long
   
On Error GoTo ErrHnd

10    If TypeOf mobjGUI.pctlData Is TextBox Then
         '''If not custom length - Called only of current data control is text box
20       If (mobjGUI.pctlData.MaxLength = 0) Then
40          Select Case mobjData.plType
            Case lDTYP_Integer, lDTYP_Long, lDTYP_Double
               '''If Maximum value is specified
100            If Len(gobjLibStr.LibStr_sEmptyIfNull(mobjData.pvMax, True)) > 0 Then
                  '''Length of the maximum value allowed input
120               lMaxLen = Len(Format$(mobjData.pvMax, mobjGUI.psFormat))
                  If lMaxLen = 0 Then
                     lMaxLen = 1
                  End If
               Else
                  lMaxLen = 20   '''Some limit
               End If
               '''If Minimum value is specified
200            If Len(gobjLibStr.LibStr_sEmptyIfNull(mobjData.pvMin, True)) > 0 Then
                  '''Length of the minimum value allowed input
220               lLenMinVal = Len(Format$(mobjData.pvMin, mobjGUI.psFormat))
                  If lLenMinVal > lMaxLen Then
                     lMaxLen = lLenMinVal
                  End If
               End If
            Case lDTYP_Date
               '''2(day) + 9(lenght of September - longest month name) + 4(year) + 2(spaces)
               lMaxLen = 17
            Case lDTYP_DateTime
               '''2(day) + 9(lenght of September - longest month name) + 4(year) + 8(time) + 3(spaces)
               lMaxLen = 26
            Case lDTYP_Time
               '''Hh:Nn:Ss
               lMaxLen = 8
            Case lDTYP_USPhone
               lMaxLen = 21
            Case lDTYP_InterPhone
               lMaxLen = 25
            Case Else   '''Text type entires
               lMaxLen = gobjLibValid.LibValid_dVal(gobjLibStr.LibStr_sEmptyIfNull(mobjData.pvMax, True))
            End Select
            
400         mobjGUI.pctlData.MaxLength = lMaxLen

         End If
      ElseIf TypeOf mobjGUI.pctlData Is DTPicker Then
500      If IsDate(mobjData.pvMax) Then
540         mobjGUI.pctlData.MaxDate = mobjData.pvMax
         End If
600      If IsDate(mobjData.pvMin) Then
640         mobjGUI.pctlData.MinDate = mobjData.pvMin
         End If
      End If
      
999   M_bSetMinMax = True
   
Exit Function
ErrHnd:
   M_bSetMinMax = False
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

Public Function P_bInit _
      (Optional ByVal lReset As pnumRV_ResetValue = lRV_Current, _
       Optional ByVal vNewValue As Variant _
      ) As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' lReset
'''   Indicates which value will be used as initial value. Defaults to lRV_Current
''' vNewValue
'''   Optional new value of the Item. Defaults to empty
'''   Ignored if lReset is not lRV_New
'''
'''Description:   This procedure indicates that client have completed setting up object properties.
'''               Apply specified properties.
'''               Indicate that current value is an initial value.
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - True
'''               Else           - False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bInit"

On Error GoTo ErrHnd

      mbInit = False
20    Call M_bSetDefaultFormat
40    If mobjItem.P_bInit(lReset, vNewValue) Then
         '''If control is set
100      If Not mobjGUI.pctlData Is Nothing Then
300         Call M_bSetMinMax
400         Call mobjGUI.P_bRefresh
         End If
500      mbInit = True
      End If
      
999   P_bInit = mbInit
   
Exit Function
ErrHnd:
   P_bInit = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function


Public Property Get pobjData() As mclsLibFieldData
   Set pobjData = mobjData
End Property
Public Property Get pobjGUI() As mclsLibFieldGUI
   Set pobjGUI = mobjGUI
End Property

Public Property Get pobjItem() As mclsLibItem
   Set pobjItem = mobjItem
End Property


Private Sub Class_Initialize()
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
'''
'''Description:   Initialize local objects.
'''               Share Item object between GUI and data objects
'''               to allow for consistent reference of the common properties
'''               and to allow synchronization through its events
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Class_Initialize"

   Dim bResult          As Boolean
   Dim vVal             As Variant
   
On Error GoTo ErrHnd
   
10 Set mobjItem = New mclsLibItem
20 Set mobjData = New mclsLibFieldData
30 Set mobjGUI = New mclsLibFieldGUI

40 Set mobjData.fobjItem = mobjItem
50 Set mobjGUI.fobjItem = mobjItem

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Sub

Private Sub Class_Terminate()
   Set mobjItem = Nothing
   Set mobjGUI = Nothing
   Set mobjData = Nothing
End Sub

Private Sub mobjGUI_pevCtlDataSet()
   '''If change of initial control
   If mbInit Then
      Call P_bInit   '''Initialize it
   End If
End Sub


