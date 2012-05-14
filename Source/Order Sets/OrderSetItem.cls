VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrderSetItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''************************************************************************
'''Module:              clsOrderSetItem/OrderSetItem.cls
'''
'''Description:         Order Item and associated additional properties
'''                     required for Order Sets
'''
'''Modules Required:    gobjShared
'''
'''Author         :     Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME  As String = "clsOrderSetItem."

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''


Public pbSelected       As Boolean
Public pbModify         As Boolean

Private mlType          As enumItemType
Private msKey           As String
Private msIcon          As String
Private mbQueueItem     As Boolean
Private msSortName      As String
Private msName          As String
Private msDetails       As String
Private mlHLOrderType   As hlORDERTYPECODE

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mobjOrderItem   As Object


Public Function Export_GEMS_Item() As String
        
    Dim colReturn As Collection
    Set colReturn = New Collection
    
    If TypeOf pobjOrderItem Is GEMS_OE.RxType Then
       colReturn.Add "Med"
       colReturn.Add pobjOrderItem.Comment
       colReturn.Add pobjOrderItem.sigCode
       colReturn.Add pobjOrderItem.Quantity
       colReturn.Add pobjOrderItem.Refills
       colReturn.Add pobjOrderItem.PharmacyNameSent
       colReturn.Add pobjOrderItem.PharmacyIenSent
       colReturn.Add pobjOrderItem.ChildCap
       colReturn.Add pobjOrderItem.OrderIen
       colReturn.Add pobjOrderItem.OrderName
    ElseIf TypeOf pobjOrderItem Is GEMS_OE.LabType Then
       colReturn.Add "Lab"
       colReturn.Add pobjOrderItem.OrderIen
       colReturn.Add pobjOrderItem.OrderName
       colReturn.Add pobjOrderItem.Comment
       colReturn.Add pobjOrderItem.CollectMethod
       colReturn.Add pobjOrderItem.ProcessPriority
       colReturn.Add pobjOrderItem.LabCollectPriority
       colReturn.Add pobjOrderItem.Specimen.Name
       colReturn.Add pobjOrderItem.Specimen.IEN
       colReturn.Add pobjOrderItem.Specimen.Topography
       colReturn.Add pobjOrderItem.Specimen.TopographyIen
       colReturn.Add pobjOrderItem.Specimen.AnswerCode
       colReturn.Add pobjOrderItem.Specimen.AnswerText
       colReturn.Add pobjOrderItem.Specimen.CapColorOrContainer
       colReturn.Add pobjOrderItem.Timing.OnexOrContinued
       colReturn.Add pobjOrderItem.Timing.AdminTime
       colReturn.Add pobjOrderItem.Timing.Duration
       colReturn.Add pobjOrderItem.Timing.Frequency
       colReturn.Add pobjOrderItem.Timing.StartDate
    Else
       colReturn.Add "Rad"
       colReturn.Add pobjOrderItem.OrderIen
       colReturn.Add pobjOrderItem.OrderName
       colReturn.Add pobjOrderItem.Comment
       colReturn.Add pobjOrderItem.ClinicalImpressionText
       colReturn.Add pobjOrderItem.ProcessPriority
       colReturn.Add pobjOrderItem.TransportMode
       colReturn.Add pobjOrderItem.WetReadingWanted
       colReturn.Add pobjOrderItem.Timing.OnexOrContinued
       colReturn.Add pobjOrderItem.Timing.AdminTime
       colReturn.Add pobjOrderItem.Timing.Duration
       colReturn.Add pobjOrderItem.Timing.Frequency
       colReturn.Add pobjOrderItem.Timing.StartDate
    End If
    
    Export_GEMS_Item = SerializeCollection(colReturn, gsRecordDelim)

End Function


Public Function ExportItem() As String
        
    Dim colReturn As Collection
    Set colReturn = New Collection
    
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
    ExportItem = Export_GEMS_Item
    Exit Function

End Function

Public Function P_bVerify _
(objOESMan As Object, ByVal bModify, ByRef bCancel As Boolean, ByRef sErrMsg As String) _
 As Boolean
'''************************************************************************
'''Called By:     P_bSubmit
'''
'''Arguments:
''' objOESMan     Reference to Order Entry Session Manager
'''                  Assumed to be available
''' bModify       Indicates whether Item need to submited with option to modify
''' bCancel       Indicates whether Item was canceled. To be returned byref
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Prepare Item for submission - make sure all required properties are populated and valid
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If order was successfully verified
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bVerify"

   Dim lLine            As Long
   Dim sImpression      As String
   Dim sPrompt          As String
   
On Error GoTo ErrHnd
                                                                                                lLine = 10
   sErrMsg = ""
                                                                                                lLine = 30
   With objOESMan
      '''Initialize Template
      Select Case mlType
         Case itMed
                                                                                                lLine = 50
            Set .RXTablet = mobjOrderItem.CopyTablet
         Case itLab
                                                                                                lLine = 60
            '--- SCR 63091   Sherry Wang   12/22/2004
            Set .LabTablet = mobjOrderItem
         Case itRad
                                                                                                lLine = 70
            '--- SCR 63091   Sherry Wang   12/22/2004
            Set .RadTablet = mobjOrderItem
         Case Else
                                                                                                lLine = 80
            sErrMsg = LibStr_sBuildSentence(gsINVALID, gsORDER, gsTYPE, ":", mlType)
      End Select
                                                                                                lLine = 90
      If Not mbQueueItem And sErrMsg = "" Then
                                                                                                lLine = 100
         '''Verify Item
         Select Case .PerformAction(mlHLOrderType, hlACTION_PRESET)
            Case hlERROR_NONE
               '''Continue
                                                                                                lLine = 110
            Case hlERROR_RX_REASON, hlERROR_USER_CANCELLED
                                                                                                lLine = 120
               bCancel = True
            Case Else
                                                                                                lLine = 130
               sErrMsg = LibStr_sBuildSentence(gsFAILED, gsTO, gsEVALUATE, gsITEM)
         End Select
      End If
                                                                                                lLine = 200
      '''If no errors and not cancel
      If sErrMsg = "" And Not bCancel Then
                                                                                                lLine = 220
         If Not mbQueueItem Then
            '''Update Properties of the Template Item - could be out of date
            Select Case mlType
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Med
               Case itMed
                    '** 67126 - Frank
                    '** 3/1/05 Added Default Start Date and Stop Date
                   .RXTablet.Timing.StartDate = Now
                   .RXTablet.Timing.StopDate = Now
                                                                                                lLine = 320
                  .RXTablet.OrderDate = Now
                                                                                                lLine = 340
                  .RXTablet.PharmacyIenSent = .RXTablet.DefaultPharmacyIen
                                                                                                lLine = 360
                  .RXTablet.PharmacyNameSent = .RXTablet.DefaultPharmacyName
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Lab
               Case itLab
                                                                                                lLine = 420
                  .LabTablet.Timing.StartDate = Now
                                                                                                lLine = 440
                  .LabTablet.Timing.AdminTime = Format(Now, "hh:mm")
                                                                                                lLine = 460
                  .LabTablet.SpecimenCollectionDate = Now
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Rad
               Case itRad
                                                                                                lLine = 520
                  .RadTablet.Timing.StartDate = Now
                                                                                                lLine = 540
                  If bModify Then
                     '''Prompt for Clinical Impression
                                                                                                lLine = 550
                     sImpression = .RadTablet.ClinicalImpressionText
                                                                                                lLine = 560
                     sPrompt = LibStr_sBuildSentence _
                              (gsCLINICAL, gsIMPRESSION, gsFOR, gsORDER) & ": " & msName
                                                                                                lLine = 570
                     sImpression = pobjMsgBox.LibMsg_sInputStr(gsORDER_SETS, sPrompt, sImpression, bCancel, True, , msDetails, True, True)
                      
                                                                                                lLine = 580
                     If Not bCancel Then
                                                                                                lLine = 590
                        .RadTablet.ClinicalImpressionText = sImpression
                     End If
                  End If
            End Select
         End If
      End If
   End With
                                                                                                lLine = 640
   If bCancel Then
                                                                                                lLine = 650
      sErrMsg = sErrMsg & " - " & LibStr_sBuildSentence(gsORDER, gsWAS, gsCANCELED)
   End If
                                                                                                lLine = 660
   P_bVerify = sErrMsg = ""
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   P_bVerify = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, _
                     gsLINE_ & lLine & ", " & gsTYPE & ": " & mlType _
                    )
End Function


Private Function M_sDetails() As String
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''
'''Description:   If order item object is set
'''                  Based on current order type
'''                  Return appropriate Details information
'''               else
'''                  Report an error
'''Note:
'''
'''Updates Globals:
'''Returns:             If succcessful
'''                        Value for Details column
'''                     else
'''                        Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "M_sDetails"
   
   Dim sResult                As String
   
On Error GoTo ErrHnd

   Select Case mlType
      Case itMed
         ''For medications return sig code, quantity and refills
         With mobjOrderItem
            sResult = .sigCode
            '** SCR 69386 -- 12/16/05 - Frank; Add Qty and Refills if not already include in SigCode
            If InStr(1, .sigCode, "#") = 0 Then sResult = sResult & " #" & .Quantity
            If InStr(1, .sigCode, "RF") = 0 Then sResult = sResult & " RF" & .Refills
            sResult = Replace(sResult, "  ", " ")   '** Remove double-space
         End With
      Case itRad, itLab
         ''For Labs and Radiographs return priority
         Select Case mobjOrderItem.ProcessPriority
            Case 1
               sResult = "Routine"
            Case 2
               sResult = "STAT"
            Case 3
               sResult = "ASAP"
            Case 4
               sResult = "NOTIFY"
            Case 5
               sResult = "PREOP"
            Case Else
               sResult = "UNKNOWN"
         End Select
         If mlType = itRad Then
            If mobjOrderItem.WetReadingWanted = hlWET_READING_YES Then
               sResult = sResult & ", Immediate Read"
            End If
         End If
         
         'scr 24598 7/30/02 Phil Crowder
         If mlType = itLab Then
            sResult = sResult & Space$(1) & mobjOrderItem.Specimen.Name
         End If
      Case Else
         '''Invalid type
         sResult = "Invalid type"
   End Select
   
   M_sDetails = sResult

Exit Function
ErrHnd:
   sResult = Err.Description
   M_sDetails = sResult
   Call LibApp_Error(Err.Number, sResult, msMODULE_NAME, sPROC_NAME, "")

End Function


Public Function P_bSetOrderItem _
      (objItem As Object, ByVal sItemKey As String, ByVal bQueueItem As Boolean, ByRef sErrMsg As String) _
       As Boolean
'''************************************************************************
'''Called by:     Any
'''
'''Arguments:
''' objItem       Order Item(type Rx, Lab or Med)
''' sItemKey      Key of the specified queue Item - used for Mtd_bRemoveOrder
''' bQueueItem    Indicates whether Item is part of the queue (else - template item)
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   If specified object is set and one of the valid type
'''                  accept it
'''               Else
'''                  Report an error
'''                  Return false
'''
'''Note:
'''Uses Globals: None
'''
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "P_bSetOrderItem"
   
   Dim sTypeName     As String
   Dim lLine         As Long
   
On Error GoTo ErrHnd
   
                                                                              lLine = 10
   sErrMsg = ""
   sTypeName = TypeName(objItem)
                                                                              lLine = 20
   If Not objItem Is Nothing Then
      '''If Key is specified
      If Not sItemKey = "" Then
         '''Determine and store type
         If TypeOf objItem Is HL7Orders.RxType Or TypeOf objItem Is GEMS_OE.RxType Then
            mlType = itMed
            msSortName = IIf(bQueueItem, "Rx", "RxQ")
            msIcon = "RX"
            mlHLOrderType = hlRX_ORDERS_CODE
         ElseIf TypeOf objItem Is HL7Orders.RadType Or TypeOf objItem Is GEMS_OE.RadType Then
            mlType = itRad
            msSortName = IIf(bQueueItem, "Rad", "RadQ")
            msIcon = "RAD"
            mlHLOrderType = hlRAD_ORDERS_CODE
         ElseIf TypeOf objItem Is HL7Orders.LabType Or TypeOf objItem Is GEMS_OE.LabType Then
            mlType = itLab
            msSortName = IIf(bQueueItem, "Lab", "LabQ")
            msIcon = "LAB"
            mlHLOrderType = hlLAB_ORDERS_CODE
         Else
            sErrMsg = "Invalid Type: " & sTypeName
         End If
      Else
         sErrMsg = "Item key is not specified, Type: " & sTypeName
      End If
   Else
      sErrMsg = "Object is not set"
   End If
   
   If sErrMsg = "" Then
                                                                              lLine = 100
      '''Store all properties of the specified object
      '--- SCR 63091   Sherry Wang   12/22/2004
      If TypeOf objItem Is HL7Orders.RxType Or TypeOf objItem Is GEMS_OE.RxType Then
        Set mobjOrderItem = objItem.CopyTablet
      Else
        Set mobjOrderItem = objItem
      End If
      msKey = sItemKey
      mbQueueItem = bQueueItem
      msName = mobjOrderItem.OrderName
      msName = Left$(Trim$(msName), 100)
      msDetails = Left$(Trim$(M_sDetails), 256)
   Else
      Call LibApp_Error(1, sErrMsg, msMODULE_NAME, sPROC_NAME, "")
   End If
   
   P_bSetOrderItem = sErrMsg = ""

Exit Function
ErrHnd:
   P_bSetOrderItem = False
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE & ": " & lLine)
   
End Function


Public Function P_bSubmit _
(objOESMan As Object, ByVal bModify, ByRef bForceModify As Boolean, _
 ByRef bCancelBatch As Boolean, ByRef sErrMsg As String, ByVal bResetOrderingProvider As Boolean _
) _
 As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' objOESMan     Reference to Order Entry Session Manager
'''                  Assumed to be available
''' bModify       Indicates whether Item need to submited with option to modify
''' bForceModify  Indicates whether Item need to be submitted again with option to modify
'''               Note: applies to non-modify Items only
''' bCancelBatch  Reference to cancel batch flag. Used to prevent submitting an item if
'''               batch was canceled during verify
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   Submit specified Item to OE for processing
'''
'''Note:
'''Warning:
'''
'''Returns:       True   - If Item was submitted
'''               False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "P_bSubmit"

   Dim lLine            As Long
   Dim bResult          As Long
   Dim bCancel          As Boolean
   
    Dim mbIsPregnant As Boolean
     Dim MedID As String, DrugName As String
     Dim IsUserCancel As Boolean, OverrideReasonID As String, OverrideComment As String, WarningInfo As String
    
    
On Error GoTo ErrHnd

                                                                                                lLine = 20
   sErrMsg = ""
                                                                                                lLine = 40
   '''<SCR 254(Pharmacy) - must use RxPOE interface for BDM pharmacy
   If (mlHLOrderType = hlRX_ORDERS_CODE) And (gobjShared.pbMTFSupportsBDM) Then
                                                                                                lLine = 48
      Call pobjMsgBox.LibMsg_lOKCancel(lMSGOK_CURRENTLY_NOT_SUPPORTED, , , "Submit Rx Order", "by RxTFC", gsORDER_SETS)
   '''>SCR 254
   Else
   With objOESMan
                                                                                                lLine = 60
      Call .Reset(mlHLOrderType)    '''Clear OE Session Manager Errors
                                                                                                lLine = 80
      '''Prepare Item for submission - make sure all required properties are populated and valid
      If P_bVerify(objOESMan, bModify, bCancel, sErrMsg) Then
                                                                                                lLine = 100
         '''If Item or Batch where not canceled
         If (Not bCancel) And (Not bCancelBatch) Then
                                                                                                lLine = 120
            If bModify Then
                                                                                                
                '1/12/05 Frank
                'SCR 64160
                If mlHLOrderType = hlRX_ORDERS_CODE Then
                    bResult = ScreenDrug(objOESMan)
                Else
                    bResult = True  '** For Rad/Lab
                End If
                    
                         
                         lLine = 130
                '** User CANCEL from the Drug Warning screen
                '** SCR 67565 - Frank - 3/12/05
                If bResult Then
                    Call .RequestToUpdateScreen(mlHLOrderType, hlACTION_NEW)
                    
                    'SCR 53311 - Send message to allow editing of order
                    Call cwobjCWComm.Message(cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION)
                                                                                     lLine = 140
                    bResult = True                                        '''Successful submission
                End If
            Else
            
               'SCR 53311 - Restablish lock on core if not modifying
               Call cwobjCWComm.Message(cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION)
                 
              'SCR 24559 7/30/03 Phil Crowder
              If mlHLOrderType = hlLAB_ORDERS_CODE Then
                 'if the particular order set item has required comments then call RequesttoUpdateScreen
                 If .LabTablet.OrderSetHasReqComments = True Then
                    Call gobjAnPParent.UnlockAnP("clsOrdSets.")
                    Call .RequestToUpdateScreen(hlLAB_ORDERS_CODE, hlACTION_NEW)
                    
                    bResult = True                                        '''Successful submission
                 End If
              End If
                                                                                                               
              'Reset the ordering provider ien if required
              If Not bResult And bResetOrderingProvider Then bResult = .PerformAction(mlHLOrderType, hlACTION_RESET_ORDERING_PROVIDER)
              
              If Not bResult Then
                                                                          
                                                                                                                
                    '1/12/05 Frank
                    'SCR
                    If mlHLOrderType = hlRX_ORDERS_CODE Then
                        '*** Screen DRUG
                        bResult = ScreenDrug(objOESMan)
                    Else
                        bResult = True '** For Rad/Lab
                    End If
                                                                                                
                    '*** 1/27/05 Frank
                    '*** Move Select statement below from inside the IF-Statement above
                    '*** So that Rad/Lab will work as before
                    
                    If bResult Then
                                                                                                    lLine = 150
                        Select Case .PerformAction(mlHLOrderType, hlACTION_NEW)
                           Case hlERROR_NONE
                                                                                                         lLine = 152
                              bResult = True                                  '''Successful submission
        
                                                                                                         lLine = 160
                           Case hlERROR_SOFT_ERROR
                                                                                                         lLine = 170
                              bForceModify = True
                           Case Else
                                                                                                         lLine = 180
                              sErrMsg = LibStr_sBuildSentence(gsFAILED, gsTO, gsCREATE, gsNEW, gsORDER)
                        End Select
                               
                    End If
                End If
            End If
         End If
      End If
   End With
   End If
   
   lLine = 200
   
   P_bSubmit = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   P_bSubmit = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, _
                     gsLINE_ & lLine & ", " & gsTYPE & ": " & mlType _
                    )
    Exit Function
    Resume
End Function

Private Function ScreenDrug(objOESMan As Object) As Boolean

    '*** Screen DRUG
    Dim mbIsPregnant As Boolean
    Dim MedID As String, DrugName As String
    Dim IsUserCancel As Boolean, OverrideReasonID As String, OverrideComment As String, WarningInfo As String

    ScreenDrug = False
    
    If mlHLOrderType = hlRX_ORDERS_CODE Then
        With objOESMan
            MedID = .RXTablet.OrderIen
            mbIsPregnant = gobjAnPParent.IsPregnant
            If .PerformScreenMed(MedID, DrugName, IsUserCancel, OverrideReasonID, OverrideComment, WarningInfo, mbIsPregnant) = hlERROR_NONE Then
                If IsUserCancel = True Then
                    '*** When user Cancel drug with warning,
                    '*** do not add the drug
                    ScreenDrug = False
                    Exit Function
                Else
                    .RXTablet.OverrideReasonID = OverrideReasonID
                    .RXTablet.OverrideComment = OverrideComment
                    .RXTablet.WarningInfo = WarningInfo
                    ScreenDrug = True
                End If
            End If
        End With
    End If
    
End Function

Public Property Get pbQueueItem() As Boolean
   pbQueueItem = mbQueueItem
End Property

Public Property Get plHLOrderType() As hlORDERTYPECODE
   plHLOrderType = mlHLOrderType
   
End Property

Public Property Get plType() As enumItemType
   plType = mlType
End Property

Public Property Get pobjOrderItem() As Object
   Set pobjOrderItem = mobjOrderItem
End Property

Public Property Get psDetails() As String
   psDetails = msDetails
End Property

Public Property Get psIcon() As String
   psIcon = msIcon
End Property

Public Property Get psKey() As String
   psKey = msKey
End Property

Public Property Get psName() As String
   psName = msName
End Property

Public Property Get psSortName() As String
   psSortName = msSortName
End Property

Private Sub Class_Initialize()
   mlType = -99 '''Default to invaid type
End Sub

Private Sub Class_Terminate()
   Set mobjOrderItem = Nothing
End Sub

