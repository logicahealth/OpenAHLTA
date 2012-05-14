VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           mclsLibItem/LibItem.cls - Generic Item class
'''
'''Description:      Generic Item
'''                  Common properties methods and events of the common data element
'''Note:
'''
'''Modules Required:
'''
'''
'''Author         :  Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME   As String = "mclsLibItem."

'''Item Status
Public Enum pnumVST_ValueStatus
   lVST_Init            ''' Indicates whether Value is initialized
   lVST_Modified        ''' Indicates whether Value was modified
   lVST_New             ''' Indicates whether Value contents is new.
   lVST_Empty           ''' Indicates whether Value is empty
   lVST_Deleted         ''' Indicates whether Value is deleted
   lVST_zCount
End Enum
'''Possible Values for P_bInit
Public Enum pnumRV_ResetValue
   lRV_Empty
   lRV_Default
   lRV_Initial
   lRV_Current
   lRV_Previous
   lRV_New
End Enum
''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private msKey                 As String
Private mbKeyIsLocked         As Boolean

Private msName                As String
Private mvValue               As Variant
Private mvValueInitial        As Variant
Private mvValuePrevious       As Variant
Public pvValueDefault         As Variant

Private mbFreezeEvents        As Boolean
Private mbInit                As Boolean  '''Indicates whether initial value was set

Public psDesc                 As String
Public pbActive               As Boolean
Private mbEditable            As Boolean  '''Indicates whether value can be chaged. Defaults to True. Will cause an error in Let pvValue if false.
Private mbEditableIsLocked    As Boolean

Public pbSelected             As Boolean  '''Reserved for client
Public pvTag                  As Variant  '''Reserved for client
Public plAllowComments        As pnumYNA_YesNoAny  '''Defaults to NO
Public plMinLenComments       As Long
Public psComments             As String

''''''''''''''' Events Section '''''''''''''''''''''''''''''''''''''''''

Public Event pevValueBeforeChange(ByRef vNewValue As Variant, ByRef bCancel As Boolean)
Public Event pevValueChanged()
Public Event pevNameChanged()
Public Event pevEditableChanged(ByVal bNewVal As Boolean)
Public Event pevBeforeGetValue()
Friend Property Let fbFreezeEvents(ByVal bVal As Boolean)
   mbFreezeEvents = bVal
End Property


Public Sub P_LockKey()
   mbKeyIsLocked = True
End Sub

Public Property Get pbEditable() As Boolean
   pbEditable = mbEditable
End Property

Public Property Let pbEditable(ByVal bVal As Boolean)
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' bVal
'''   Requested new value
'''
'''Description:   Update editable property. Inform client about the change
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pbEditable"
   
   Static bIn_static    As Boolean

On Error GoTo ErrHnd

20 If (Not bIn_static) Then
      If Not mbEditableIsLocked Then
100      bIn_static = True
120      mbEditable = bVal
140      RaiseEvent pevEditableChanged(mbEditable)
      Else
         '''Raise error
200      Call gobjLibApp.LibApp_Error _
         (1, gobjLibStr.LibStr_sBuildSentence(gsEDIT, gsIS, gsLOCKED), _
          msMODULE_NAME, sPROC_NAME, psName, True, False, False _
         )
      End If
220   bIn_static = False
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True, False, False)
   bIn_static = False
   
End Property

Public Property Let pbEditableIsLocked(ByVal bVal As Boolean)
   mbEditableIsLocked = bVal
End Property

Public Property Let psKey(ByVal sVal As String)
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' sVal
'''
'''Description:      If Key value is not protected accept specified value
'''                  Else report an error
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let psKey"
   
   If Not mbKeyIsLocked Then
      msKey = sVal
   Else
      '''Raise error
      Call gobjLibApp.LibApp_Error _
      (1, gobjLibStr.LibStr_sBuildSentence(gsKEY, gsIS, gsLOCKED), _
       msMODULE_NAME, sPROC_NAME, psName, True, False, False _
      )
   End If
   
End Property

Public Property Get psKey() As String
   psKey = msKey
End Property

Public Property Let psName(ByVal sVal As String)
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' vVal
'''   Requested new name(caption,label) of the Item
'''
'''Description:   Update current item name
'''               Inform client about the change
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let psName"
   
   Static bIn_static    As Boolean

On Error GoTo ErrHnd
   
   If Not bIn_static Then
      bIn_static = True
20    msName = sVal
      If Not mbFreezeEvents Then
100      RaiseEvent pevNameChanged
      End If
      bIn_static = False
   End If

Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error _
       (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True, False, False)
   bIn_static = False
   
End Property

Public Property Get psName() As String
   psName = msName
End Property

Public Property Get pbStatus(ByVal lAttribute As pnumVST_ValueStatus) As Boolean
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' lAttribute
'''   Indicates which one of the status attributes is requested e.g. New, Deleted, Modified...
'''
'''Description:      Evaluate and return current status of the value
'''                  corresponing to requested attribute
'''Note:
'''
'''Warning:          Will cause en error if P_bInit was not called and requested attribute
'''                  requires object to be initilized ( not lVST_Init or lVST_Empty)
'''
'''Updates Globals:
'''
'''Returns:          If successful  - Status of the Value
'''                  Else           - Causes an error
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pbStatus"
   
   Dim bResult          As Boolean
   
   '''If requested attribute requires this object to be initialized
   If Not lAttribute = lVST_Init Then
      Call gobjLibApp.LibApp_sInitErr(True, mbInit) '''Error out if not initialized(P_bInit was not called)
   End If
   
   With gobjLibStr
      Select Case lAttribute
         Case lVST_Init
            bResult = mbInit
         Case lVST_Modified
            bResult = Not _
                     (gobjLibValid.LibValid_vCastType(mvValueInitial) = _
                      gobjLibValid.LibValid_vCastType(mvValue) _
                     )
         Case lVST_New
            '''New if initial value was empty and current is not
            bResult = (.LibStr_sEmptyIfNull(mvValueInitial) = "") And _
                      (Not .LibStr_sEmptyIfNull(mvValue) = "")
         Case lVST_Empty
            bResult = Len(.LibStr_sEmptyIfNull(mvValue)) = 0
         Case lVST_Deleted
            '''Deleted if initial value was not empty and current is
            bResult = (Not .LibStr_sEmptyIfNull(mvValueInitial) = "") And _
                      .LibStr_sEmptyIfNull(mvValue) = ""
         Case Else
            '''Raise error
            Call gobjLibApp.LibApp_Error _
            (1, gobjLibStr.LibStr_sBuildSentence(gsUNDEFINED, gsATTRIBUTE), _
             msMODULE_NAME, sPROC_NAME, psName, True, False, False _
            )
      End Select
   End With
   
   pbStatus = bResult
   
End Property

Public Property Let pvValue(ByVal vValue As Variant)
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' vValue
'''   Requested new value
'''
'''Description:   If Item is editable
'''                  update current value
'''                  Inform client about the change
'''               else
'''                  raise error
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Let pvValue"
   
   Dim bCancel          As Boolean
   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd

   If Not bIn_static Then
10    bIn_static = True
      '''Not Non-editable value can only be reset by calling P_bInit
20    If pbEditable Then
40       If Not mbFreezeEvents Then
            '''Inform about new value and allow to modify it or cancel
60          RaiseEvent pevValueBeforeChange(vValue, bCancel)
         End If
         If Not bCancel Then
100         mvValuePrevious = mvValue
120         mvValue = vValue
            If Not mbFreezeEvents Then
200            RaiseEvent pevValueChanged
            End If
         End If
      Else
         '''Raise error
500      Call gobjLibApp.LibApp_Error _
         (1, gobjLibStr.LibStr_sBuildSentence(gsLOCKED, gsVALUE), _
          msMODULE_NAME, sPROC_NAME, psName, True, False, False _
         )
      End If
      bIn_static = False
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True, False, False)
   bIn_static = False
   
End Property

Public Property Get pvValue() As Variant
   RaiseEvent pevBeforeGetValue
   pvValue = mvValue
End Property

Public Property Get pvValueInitial() As Variant
   pvValueInitial = mvValueInitial
End Property

Public Property Get pvValuePrevious() As Variant
   pvValuePrevious = mvValuePrevious
End Property

Private Sub Class_Initialize()
   pbActive = True
   mbEditable = True
   plAllowComments = lYNA_No
   plMinLenComments = 10
End Sub


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
'''Description:      Reset initial value to current
'''                  Indicate that initial value was set - object is ready to be used
'''
'''Note:             Determines outcome of the psStatus property
'''                  vNewValue will override current value even if pbEditable is false
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:          True  - If successful
'''                  False - Else
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bInit"

On Error GoTo ErrHnd
      
20    Select Case lReset
         Case lRV_Empty
100         mvValue = Empty
         Case lRV_Default
200         mvValue = pvValueDefault
         Case lRV_Initial
300         mvValue = mvValueInitial
         Case lRV_Current
            '''Do nothing - keep current
         Case lRV_New
500         If IsMissing(vNewValue) Then
520            mvValue = Empty
            Else
540            mvValue = vNewValue
            End If
         Case lRV_Previous
600         mvValue = mvValuePrevious
         Case Else
700         Call gobjLibApp.LibApp_Error _
            (1, gobjLibMsg.LibMsg_sOKCancelText(lMSGOK_INVALID_OPTION), msMODULE_NAME, sPROC_NAME, _
             gsRESET & ": " & lReset, True _
            )
      End Select
      
800   mvValuePrevious = Empty
820   mvValueInitial = mvValue
840   mbInit = True

999   P_bInit = True
   
Exit Function
ErrHnd:
   P_bInit = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, True)

End Function

