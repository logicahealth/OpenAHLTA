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
Attribute VB_Name = "mclsLibFieldData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           clsLibFieldData/LibFieldData.cls - Data field object
'''
'''Description:      Allows to define, store and evaluate single data element
'''
'''Modules Required:
'''
'''Note:
'''
'''Warning:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME           As String = "mclsLibFieldData."

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''
                                                   
'''''' Local use
Private mlLastPopTime         As Long              '''Used to avoid to lock into lost focus event when it causes validation
Private mbClearRequest        As Boolean           '''Used to prevent triggering validation when value is set in P_Clear

''''''Properties
                                                   ''' DEFAULT
Private mlType                As gnumDTYPDataType  ''' vbEmpty - Undefined Data Type
Private mbAcceptInvalid       As Boolean           ''' True - Indicates whether accept invalid value or report an error
Private mvMax                 As Variant           ''' Empty - Ignored if empty. For string represents maximum number of characters - else represents maximum value.
Private mvMin                 As Variant           ''' Empty - Ignored if empty. For string represents minimum number of characters - else represents minimum value.
Private mbRequired            As Boolean           ''' False - Indicates whether Empty(Null) value shoud cause validation error
Public pvTag                  As Variant           ''' Empty - Allows client to associate any value. Reserved for client
Public pbPopMsg               As Boolean           ''' True - Indicates whether to Display OK Message for invalid values



''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

'''<DSN may not be needed
Public Event pevValidate(ByVal vValue As Variant, ByRef sErrMsg As String, ByRef bInvalid As Boolean)   '''Allows for custom validation prior to standard validation. If caller intercepts and event and changes bInvalid to true standard validation will be canceled and P_bValid will return false

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private WithEvents mobjItem   As mclsLibItem
Attribute mobjItem.VB_VarHelpID = -1


Private Sub M_AssertValid()
'''************************************************************************
'''Called By:     mobjItem_pevValueChanged
'''               Let pbAcceptInvalid
'''               Let plType
'''               Let pvMax
'''               Let pvMin
'''Arguments
'''
'''Description:   If Accept Invalid not allowed
'''                  Evaluate current value
'''                  If invalid attempt to reset current value using
'''                  alternatives in the order as listed: Previous, Initial, Default
'''                  If all of the available values are invalid
'''                     Set current to empty
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_AssertValid"

   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd

   If (Not mbAcceptInvalid) And (Not bIn_static) And mobjItem.pbStatus(lVST_Init) Then
      bIn_static = True
      ''If new value is invalid
      If Not P_bValid(, , pbPopMsg) Then
         '''If previous value is valid
         If P_bValid(mobjItem.pvValuePrevious, , False) Then
            '''use previous
            mobjItem.pvValue = mobjItem.pvValuePrevious
         '''else if initial is valid
         ElseIf P_bValid(mobjItem.pvValueInitial, , False) Then
            '''use Initial
            mobjItem.pvValue = mobjItem.pvValueInitial
         '''else if default is valid
         ElseIf P_bValid(mobjItem.pvValueDefault, , False) Then
            '''use defult
            mobjItem.pvValue = mobjItem.pvValueDefault
         '''else attemt to fit into valid range
         ElseIf Not M_bFitIntoRange Then
            '''Else - this is where it ends - set to empty
            mobjItem.fbFreezeEvents = True
            mobjItem.pvValue = Empty
            mobjItem.fbFreezeEvents = False
         End If
      End If
      bIn_static = False
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   bIn_static = False

End Sub


Private Function M_bFitIntoRange() As Boolean
'''************************************************************************
'''Called By:     M_AssertValid
'''Arguments:
'''
'''Description:   If String and Maximum range is valid
'''                  Keep maximum allowed left characters
'''               If numeric or date/time and Mininum range is valid
'''                  Set to minimum value
'''               else
'''                  Can not fix it
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       True  - if was able to correct the value
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bFitIntoRange"

   Dim bResult          As Boolean
   
On Error GoTo ErrHnd

   Select Case mlType
      Case lDTYP_String
         If gobjLibValid.LibValid_bInteger(pvMax) Then
            If P_bValid(Left$(gobjLibStr.LibStr_sEmptyIfNull(pvValue), pvMax)) Then
               pvValue = Left$(gobjLibStr.LibStr_sEmptyIfNull(pvValue), pvMax)
               bResult = True
            End If
         End If
      Case lDTYP_Integer, lDTYP_Long, lDTYP_Double, lDTYP_Date, lDTYP_DateTime, lDTYP_Time
         If P_bValid(pvMin) Then
            pvValue = pvMin
            bResult = True
         End If
   End Select
   
   M_bFitIntoRange = bResult
   
Exit Function
ErrHnd:
   M_bFitIntoRange = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function


Public Sub P_Clear()
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Clear Value - set it to Empty
'''
'''Note:          This procedure overrides validation criteria
'''
'''Warning:       Value will be set to empty even if pbAcceptInvalid and/or pbRequired are true
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_Clear"

On Error GoTo ErrHnd

   If (Not mbClearRequest) Then
      mbClearRequest = True
      mobjItem.pvValue = Empty
      mbClearRequest = False
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   mbClearRequest = False

End Sub

Public Property Get pbAcceptInvalid() As Boolean
   pbAcceptInvalid = mbAcceptInvalid
End Property


Public Property Let pbAcceptInvalid(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W     Default
'''               R  Y     True
'''               If ignore validation was canceled
'''                  Re-evaluate current value
'''Note:
'''
'''Warning:       If false will be iterpreted as request to auto-correct if value is invalid
'''               Value will be set to the first valid as listed:
'''               Previous, Initial, Default, Range, Empty
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pbAcceptInvalid"

   
On Error GoTo ErrHnd

   mbAcceptInvalid = bVal
   Call M_AssertValid
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property


Public Function P_bValid _
      (Optional ByVal vTestValue As Variant, _
       Optional ByRef sErrMsg As String, _
       Optional ByVal bPopMsg As Boolean _
      ) As Boolean
'''************************************************************************
'''Called By:     M_AssertValid
'''               Let pvValue
'''Arguments:
''' vTestValue
'''   Optional value requested to be evaluated against current validation criteria
'''   If not specified current value will be evaluated
''' sErrMsg
'''   Optional error info if value is invalid or error
''' bPopMsg
'''   Indicates whether to Display OK MsgBox if invalid
'''
'''Description:   Evaluate requested value
'''
'''Note:          vTestValue will not be used to modify current value
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       True  - If current Item contains valid data
'''               False - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME              As String = "P_bValid"

   Dim bResult                   As Boolean
   Dim vVal                      As Variant
   Dim bFailedCustomValidation   As Boolean
   
On Error GoTo ErrHnd
   
20    sErrMsg = ""
100   With mobjItem
200      If IsMissing(vTestValue) Then
220         vVal = mobjItem.pvValue
         Else
240         vVal = vTestValue
         End If
300      If Len(gobjLibStr.LibStr_sEmptyIfNull(vVal)) = 0 Then
400         If mbRequired Then
420            sErrMsg = gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_ENTRY_REQUIRED)
            Else
               bResult = True
            End If
         Else
500         If gobjLibValid.LibValid_bValue(mlType, vVal, pvMin, pvMax, , sErrMsg, psCaption) Then
600            If pobjItem.plAllowComments = lYNA_Yes Then
620               bResult = gobjLibValid.LibValid_bValue _
                           (lDTYP_String, pobjItem.psComments, pobjItem.plMinLenComments, , , _
                            sErrMsg, psCaption & " (" & gsCOMMENTS & ")" _
                           )
               Else
                  bResult = True
               End If
            End If
         End If
         '''If passed standard validation
800      If bResult Then
            '''Allow client to perform custom validation
820         RaiseEvent pevValidate(vVal, sErrMsg, bFailedCustomValidation)
            bResult = Not bFailedCustomValidation
         End If
900      If (Not bResult) And bPopMsg And (Abs(Timer - mlLastPopTime) > 2) Then
920         mlLastPopTime = Timer
940         Call gobjLibMsg.LibMsg_lOKCancel(lMSGOK_INVALID_ENTRY, , , psCaption, sErrMsg)
         End If
      End With
   
999   P_bValid = bResult
   
Exit Function
ErrHnd:
   P_bValid = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error _
        (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl & vbCrLf & psCaption)

End Function


Public Property Let pbRequired(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' bVal
'''   New value for the property
'''
'''Description:   R  W     Default
'''               Y  Y     True
'''               Indicates whether entry is required.
'''               If true - Empty(Null) value will cause validation error
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   mbRequired = bVal
End Property

Public Property Get pbRequired() As Boolean
   pbRequired = mbRequired
End Property

Public Property Let plType(ByVal lDataType As gnumDTYPDataType)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lDataType
'''   New data type
'''
'''Description:   R  W     Default
'''               R  Y     vbEmpty
'''               Evaluate specified type
'''               If valid
'''                  Accept it else report and raise error
'''                  Re-evaluate current value
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let plType"
   
   Dim sErrMsg          As String
   
On Error GoTo ErrHnd

   If gobjLibValid.LibValid_bTypeIsSupported(lDataType, sErrMsg) Then
      mlType = lDataType
      Call M_AssertValid
   Else
      Call gobjLibApp.LibApp_Error(1, sErrMsg, msMODULE_NAME, sPROC_NAME, "", True)
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   
End Property

Public Property Get plType() As gnumDTYPDataType
   plType = mlType
End Property

Public Property Get pobjItem() As mclsLibItem
   Set pobjItem = mobjItem
End Property
Friend Property Set fobjItem(objItem As mclsLibItem)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W     Default
'''               R  Y(f)  New
'''
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Set fobjItem"

On Error GoTo ErrHnd

   Set mobjItem = objItem
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Property

Public Property Get psCaption() As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W     Default
'''               R  Y     Empty string
'''
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - label of the
'''               else           - empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psCaption"

On Error GoTo ErrHnd

   psCaption = mobjItem.psName
   
Exit Property
ErrHnd:
   psCaption = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Property

Public Property Let psCaption(sVal As String)
   mobjItem.psName = sVal
End Property

Public Property Let pvMax(ByVal vVal As Variant)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' bVal
'''   New value for the property
'''
'''Description:   R  W     Default
'''               Y  Y     Empty
'''               For string types represents maximum number of characters
'''               for US Phone represents length of Phone Extension. Zero - Interpreted as Extension is not allowed
'''               For other types represents maximum value.
'''
'''Note:          Ignored if empty.
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   mvMax = vVal
   Call M_AssertValid
End Property

Public Property Get pvMax() As Variant
   pvMax = mvMax
End Property

Public Property Get pvMin() As Variant
   pvMin = mvMin
End Property

Public Property Let pvMin(ByVal vVal As Variant)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' bVal
'''   New value for the property
'''
'''Description:   R  W     Default
'''               Y  Y     Empty
'''               For string types represents minimum number of characters
'''               For US Phone represents mlength of Phone Extension.
'''                  Zero - Interpreted as Extension is optional.
'''                  Greater then zero indicates that Extension is required
'''               For other types represents minimum value
'''
'''Note:          Ignored if empty
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   mvMin = vVal
   Call M_AssertValid

End Property

Public Property Let pvValue(vVal As Variant)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   R  W     Default
'''               R  Y     Empty
'''               If valid value is specified or object is in (ignore validation) state
'''                  Accept new value
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pvValue"

   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd

   If Not bIn_static Then
      bIn_static = True
      If mbAcceptInvalid Or mbClearRequest Then
         mobjItem.pvValue = vVal
      ElseIf P_bValid(vVal, , pbPopMsg) Then
         mobjItem.pvValue = vVal
      End If
      bIn_static = False
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   bIn_static = False
   
End Property

Public Property Get pvValue() As Variant
   pvValue = mobjItem.pvValue
End Property


Public Property Get pvValueMatchType(Optional ByVal vVal As Variant) As Variant
'''************************************************************************
'''Called By:
'''Arguments:
''' vVal
'''   Optional value requested to be casted and returned. If not specified caurrent value will be used
'''
'''Description:   Return value casting its data type (VarType) to match current type
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  current value in current type
'''               else
'''                  Null
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pvValueMatchType"

   Dim vResult          As Variant
   
On Error GoTo ErrHnd

      '''If value is not specified
20    If IsMissing(vVal) Then
         '''Use current
40       vVal = pvValue
      End If
      
      Select Case mlType
         Case lDTYP_String, lDTYP_Path, lDTYP_NetworkPath, lDTYP_USPhone, lDTYP_InterPhone
100         vResult = gobjLibValid.LibValid_vCastType(vVal, vbString, vbString)
         Case lDTYP_Date, lDTYP_Time, lDTYP_DateTime
200         vResult = gobjLibValid.LibValid_vCastType(vVal, vbDate, vbDate)
         Case lDTYP_Boolean
300         vResult = gobjLibValid.LibValid_bEvaluateBoolean(vVal)
         Case Else
400         vResult = gobjLibValid.LibValid_vCastType(vVal, mlType, mlType)
      End Select
   
999   pvValueMatchType = vResult
   
Exit Property
ErrHnd:
   pvValueMatchType = Null
   Call gobjLibApp.LibApp_Error _
       (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Property

Private Sub Class_Initialize()
   Set mobjItem = New mclsLibItem
   mbAcceptInvalid = True
   pbPopMsg = True
   
End Sub


Private Sub Class_Terminate()
   Set mobjItem = Nothing
End Sub



Private Sub mobjItem_pevValueBeforeChange(vNewValue As Variant, bCancel As Boolean)
'''************************************************************************
'''Called By:
'''
'''Arguments
'''
'''Description:   Evaluate new value and if neccessery correct or reject it
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "mobjItem_pevValueChanged"
   
   Dim sPhone           As String
   
On Error GoTo ErrHnd

   If Not mbAcceptInvalid Then
20    bCancel = Not P_bValid(vNewValue, , pbPopMsg)
   End If
   If Not bCancel Then
100   If mlType = lDTYP_USPhone Then
         '''Make sure that value is properly formated
120      If gobjLibValid.LibValid_bUSPhone(vNewValue, pvMin, pvMax, , psCaption, sPhone) Then
200         vNewValue = sPhone
         End If
      End If
   End If

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
        (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Sub

Private Sub mobjItem_pevValueChanged()
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
'''
'''Description:   If new value is invalid and Accept Invalid not allowed
'''                  Attempt to restore valid value or set it to empty
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "mobjItem_pevValueChanged"
   
On Error GoTo ErrHnd

   Call M_AssertValid
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Sub


