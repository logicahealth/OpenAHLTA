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
Attribute VB_Name = "FormCHCSII_Labels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<< SCR #43099; 43107; 43111; 43114; 43120;   Developer: Brian Mowbray 11/06/2003 04:33 PM
Option Explicit

Private oDom As DOMDocument
Private m_Form As Form
Private ms_TabLabels As String
Private mo_PatientDemographics As PatientDemographics
Private mo_PatientRec As ADODB.Recordset
Private ms_PatientUnitNumber As String
Private mdic_Report As Dictionary
Private mdic_DBField As Dictionary
Private mdic_ListDBField As Dictionary
Private mdic_ListIdxField As Dictionary
Private mdic_CtrlLock As Dictionary
Private mi_PrintOption As Integer
Private mo_Crystal As Crystal
Private mo_Report As CRAXDRT.Report
Private moCollRecSet As Collection
Private moCollFormulas As Collection
Private mo_Rst As ADODB.Recordset
Private mo_BedInfoRst As ADODB.Recordset
Private ms_FormCaption As String
Private ms_DateFormat As String
Private ms_DateTimeFormat As String
Private ms_ClinicName As String
Private ms_ClinicAddress As String
Private ms_ClinicPhone As String
Private ml_Tab As Long
Private ml_LastTab As Long
Private mb_Loading As Boolean
Private ml_StartingTab As LabelTabs
Private mb_OrderManager As Boolean

Private ml_HeaderFontSize As Long
Private ml_DetailFontSize As Long

Private mOrderType As String
Private mOrderId   As String
Private firstTime  As Boolean

Private mNewOrderID As String
Private mlSelectedIndex As Long
Private mUser As ICHCSII_User

Private mCreateNewRx As Boolean


Private Enum enTabs
  tab0 = 0
  Tab1 = 1
  tab2 = 2
  tab3 = 3
  tab4 = 4
  Tab5 = 5
End Enum


Public Property Set User(Data As Object)
    Set mUser = Data
End Property
Public Property Get User() As Object
    Set User = mUser
End Property

Public Property Let NewOrderID(Data As String)
    mNewOrderID = Data
End Property
Public Property Get NewOrderID() As String
    NewOrderID = mNewOrderID
End Property

'--- SCR 72008  Sherry Wang  7/6/2005
Public Property Let SelectedIndex(Data As Long)
    mlSelectedIndex = Data
End Property
Public Property Get SelectedIndex() As Long
    SelectedIndex = mlSelectedIndex
End Property

Public Sub Display()
  On Error Resume Next
  '<< Begin:SCR #45121;   Developer: Brian Mowbray 01/22/2004 04:46 PM
  cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT
  m_Form.Show vbModal
  cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT
'>> End: SCR #45121;
End Sub


' Module Description: Used to get priority values. These value are hardcoded in the A&P Lab module.
' This is the only way to get the context ot each value
Private Function GetPriValue(lPri As Long) As String

  On Error GoTo Func_Err
  
  Select Case lPri
  Case 0
    GetPriValue = vbNullString
  Case 1
    GetPriValue = "ROUTINE"
  Case 2
    GetPriValue = "STAT"
  Case 3
    GetPriValue = "ASAP"
  Case 4
    GetPriValue = "NOTIFY"
  Case 5
    GetPriValue = "PREOP"
  Case Else
    GetPriValue = "UNDEFINED"
  End Select

Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.GetPriValue", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug

End Function

'This is used to retrieve the Bed Location based on the bed number retrieved by the GetBedInfo method
Private Function GetBedLocation(lWardNCID As Long) As String
  Dim rst As ADODB.Recordset
  Dim sSQL As String
  On Error GoTo Func_Err
  
  Set rst = New ADODB.Recordset
  
  sSQL = "SELECT WARD_LOCATION_NAME FROM WARD_LOCATION WHERE WARD_LOCATION_NCID = (" & _
          "SELECT WARD_LOCATION_NCID FROM WARD WHERE WARD_NCID = " & lWardNCID & ")"
  
  Set rst = goDBDAS.OpenRecordset(sSQL)

  'Return false and exit out if no data is found
  If Not rst Is Nothing Then
    If rst.BOF And rst.EOF Then
      GetBedLocation = vbNullString
      GoTo Func_Exit
    End If
  Else
    GetBedLocation = vbNullString
    GoTo Func_Exit
  End If
  
  GetBedLocation = rst.Fields(0).Value & vbNullString


Func_Exit:
  On Error Resume Next
  Set rst = Nothing
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.GetBedLocation", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug

End Function
'Used to initially retrieve Inpatient information needed to populate the Bed # and Location.
'This should only get called once when the form is loaded
Private Function GetBedInfo() As Boolean
  Dim sSQL As String
  On Error GoTo Func_Err
  
  '<< Begin:SCR #43120;   Developer: Brian Mowbray 12/29/2003 11:21 AM
  sSQL = "SELECT WARD_NCID, ASSIGNED_BED_NUMBER FROM INPATIENT_ADMISSION " & _
         "WHERE UNIT_NUMBER = '" & ms_PatientUnitNumber & "' AND DISCHARGE_DATE Is Null"
  '>> End: SCR #43120;
         
  Set mo_BedInfoRst = goDBDAS.OpenRecordset(sSQL)

  'Return false and exit out if no data is found
  If Not mo_BedInfoRst Is Nothing Then
    If mo_BedInfoRst.BOF And mo_BedInfoRst.EOF Then
      GetBedInfo = False
      GoTo Func_Exit
    Else
      GetBedInfo = True
      GoTo Func_Exit
    End If
  Else
    GetBedInfo = True
    GoTo Func_Exit
  End If

Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.GetBedInfo", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug

End Function

Public Function initLabels(Optional bInitialLoad As Boolean = True, _
                           Optional CallingFromInptOrdrMgt As Boolean = False, _
                           Optional CreateNewRx As Boolean = False) As Boolean
  
  Dim i As Integer
  Dim iCurTab As Integer
  
  Dim bResetList As Boolean
  Dim PhysicianNCID As String
  Dim PhysicianName As String
  Dim DaysOfSupply As String
  Dim Dose As String
  Dim Dosage_Form As String

  On Error GoTo Func_Err
  
  Screen.MousePointer = vbHourglass
  
  initLabels = False
  mb_Loading = True
  
  
  If bInitialLoad Then
  
    mCreateNewRx = CreateNewRx
  
    Set m_Form = frmCHCSII_Labels
    
    If m_Form.cmdNewRx.Visible = False Then
          
          If Len(m_Form.txtRXOP_Medication.Text) > 0 Then
              If MsgBox("You have an unsaved med order.  Do you want to submit this new order?", vbQuestion + vbYesNo, "Unsaved Med Order") = vbYes Then
                  Call m_Form.SaveMedOrder
               Else
                  Exit Function
               End If
          End If
    End If
  
    'SCR 173421
    m_Form.cmdNewRx.Visible = CreateNewRx
    
    '*** Remember Default Printer
    m_Form.DefaultPrinter = GetDefaultPrinterDeviceName
    
    If bInitialLoad Then
        gCallingFromInptOrdrMgt = CallingFromInptOrdrMgt
        If CallingFromInptOrdrMgt Then
            If CreateNewRx Then
                mOrderId = ""
                gOriginalMedOrderID = ""
            Else
                gOriginalMedOrderID = mOrderId
                gOrderManagerSelectedMedOrderID = mOrderId
            End If
        Else
            gOriginalMedOrderID = ""
 
        End If
    End If
    m_Form.IsEditable = True
    
    Load m_Form
    
    If mOrderType = "MED" Then m_Form.OrigionalMedOrderID = mOrderId
    
    Set m_Form.LabelObject = Me
    
    'This sets up the form tabs based on the XML file and sets
    'two module level variables for the form caption and the date format
    If Not SetupTabLabels Then GoTo Func_Exit
  
    'Loads Patient Inpatient Bed Info
    GetBedInfo
  Else
    m_Form.IsEditable = False
  End If
  
  ml_LastTab = ml_Tab
  
  bResetList = (bInitialLoad) Or (ml_Tab <> m_Form.TabOne1.CurrTab)

  'Check if order is IV when current tab is either Med or IV regardless of where labels is called from
  If m_Form.TabOne1.CurrTab = 3 Or m_Form.TabOne1.CurrTab = 4 Then
    If m_Form.UserNotClickTab Then
        m_Form.UserNotClickTab = True
        If m_Form.MedDefaultRouteIsIV(m_Form.lstOrders.Text, "") Then
            m_Form.TabOne1.CurrTab = 3
        Else
            m_Form.TabOne1.CurrTab = 4
            m_Form.cmdPhysicianNameSearch.Visible = False
            m_Form.cmdDrugSearch.Visible = False
            End If
        End If
  End If
  
  ml_Tab = m_Form.TabOne1.CurrTab
  
  'Sets the form caption from the XML config file
  m_Form.Caption = ms_FormCaption
  
  'Loads Patient Data
  LoadPatientData
  
  If goDBDAS Is Nothing Then Set goDBDAS = goDBConn.CHCSII_DAS(Auto)
  'Loads the controls based on cached data and date from the corresponding domain
  LoadControls
  
  'Sets controls to be enabled and or locked
  SetCtrlPermission

  
  If bResetList Then
  'Save any unlocked tab data
    If Not bInitialLoad Then SaveLabelData
  End If
  
  'Gets the domain list data
  If GetListData() Then
    If bResetList Then
      'Save any unlocked tab data
      If Not bInitialLoad Then SaveLabelData
      'Loads the listbox on the form
      LoadList
      'Set the list selection to the first one
    End If
  Else
    'Clear the listbox because there is not data to display
    m_Form.lstOrders.Clear
    m_Form.lstOrders.AddItem "BLANK LABEL"
    
  End If
  
    'SCR 90488 4/25/06 Quan
    'SCR 131474 - LCI - 06/08/2007: Added check for workflow referenced in SCR 90488 to only be applied
    'when user is in Order Manager module only
    If mb_OrderManager Then
        '** Disable all tabs, except current tab
        iCurTab = m_Form.TabOne1.CurrTab
        
        For i = 0 To 4
            If i = iCurTab Then
                m_Form.TabOne1.TabEnabled(i) = True
            Else
                m_Form.TabOne1.TabEnabled(i) = False
            End If
        Next i
    
    
        '*** Special case, for Tab 3 and 4 (MED), both need to enable to both disable
        '*** same with Tab 1 and 2 (RAD)
        Select Case m_Form.TabOne1.CurrTab
                            
            Case 3, 4    '** For Med both IV and Med tab should be enabled
                m_Form.TabOne1.TabEnabled(3) = True
                m_Form.TabOne1.TabEnabled(4) = True
            
            Case 1, 2    '** For both Rad Film and Rad File Folder tab should be enabled
                m_Form.TabOne1.TabEnabled(1) = True
                m_Form.TabOne1.TabEnabled(2) = True
                
        End Select
       
    Else
        
        For i = 0 To 4
              m_Form.TabOne1.TabEnabled(i) = True
        Next i
            
        LoadSelected
        SetStartingListValue vbNullString
        
    End If
    
  '*** SCR 67259 - Frank
  '*** 3/31/05 Display Patient Order Type
  m_Form.lblRXIV_PatientType.Caption = ""
  m_Form.lblRXOP_PatientType.Caption = ""
  
  m_Form.lblRXIV_PatientType.Visible = False
  m_Form.lblRXOP_PatientType.Visible = False
  
  m_Form.cmdFill.Enabled = False
  m_Form.cmdDispense.Enabled = False

  Select Case ml_Tab
    Case tab0
        m_Form.dtpLS_AccessDate.Value = Now
        
        'SCR 98098 - LCI 06/16/2006
        m_Form.txtLS_AbbrevReq.Enabled = True
        m_Form.txtLS_AbbrevReq.Locked = False
        
    Case tab2
        m_Form.dtpRFL_DateTimeExam.Value = Now
    Case tab3
        '** SCR 67309 - FRANK
        '** 4/1/05 Display either Inpatient Order or Outpatient Order in Pharmacy IV tab
        m_Form.lblRXIV_PatientType.Caption = GetPatientType(m_Form.txtRXIV_OrderNumber.Text)
        m_Form.lblRXIV_PatientType.Visible = True
                
        ' 85977 - 1/24/06 - Frank
        If m_Form.lblRXIV_PatientType.Caption = "OUTPATIENT ORDER" Then
            m_Form.txtRXIV_Location.Text = "Outpatient"
        End If
        
        Dim dtDate As Date
        dtDate = Empty
        
        '*** Default Due Date today's date
        m_Form.dtpRXIV_FillDate.Value = Now
        
        '** 86571 -- Default to today's date, but intially display blank
        m_Form.dtpRXIV_DueDate.Value = Now
        
        m_Form.dtpRXIV_DueDate.Format = dtpCustom
        m_Form.dtpRXIV_DueDate.CustomFormat = " ' " 'SCR 86241 - 1/23/06 - Frank
        
        '** 86571 -- Default to today's date, but intially display blank
        m_Form.dtpRXIV_ExpDate.Value = Now
        
        m_Form.dtpRXIV_ExpDate.Format = dtpCustom
        m_Form.dtpRXIV_ExpDate.CustomFormat = " ' " 'SCR 86241 - 1/23/06 - Frank

        m_Form.lblRXIV_DispensedBy.Caption = m_Form.GetLastFill(m_Form.txtRXIV_OrderNumber.Text)
        
        m_Form.cmdFill.Enabled = Len(Trim$(m_Form.txtRXIV_Medication.Text)) > 0
        m_Form.cmdDispense.Enabled = m_Form.cmdFill.Enabled
        m_Form.cmdRxIV_History.Enabled = m_Form.cmdFill.Enabled

        m_Form.SetTabVisibility

    Case tab4
    
        '*** SCR  - Frank
        '*** 3/31/05 Allow user to update the SIG field
        m_Form.txtRXOP_SIG.Enabled = True
        m_Form.txtRXOP_SIG.Locked = False
        
        'SCR 107015 - LCI - 09/27/2006
        m_Form.txtRXOP_OverrideReason.Enabled = False
        m_Form.txtRXOP_OverrideComment.Enabled = False

       '*** SCR 67259 - Frank
       '*** 3/31/05 Display Patient Order Type
        m_Form.lblRXOP_PatientType.Caption = GetPatientType(m_Form.txtRXOP_Number.Text)
        m_Form.lblRXOP_PatientType.Visible = True
        
        ' 85977 - 1/24/06 - Frank
        If m_Form.lblRXOP_PatientType.Caption = "OUTPATIENT ORDER" Then
            m_Form.txtRXOP_Location.Text = "Outpatient"
        End If
        
        '*** SCR - FRANK
        '*** Default Dispense date today's date
        m_Form.dtpRXOP_DispenseDate.Value = Now
        m_Form.dtpRXOP_DispenseDate.Tag = ""
        m_Form.dtpRXOP_DispenseDate.Enabled = True
        
        m_Form.lblRXOP_DispensedBy.Caption = m_Form.GetLastFill(m_Form.txtRXOP_Number.Text)
        
        m_Form.SetTabVisibility
        
        m_Form.MedOrderAction = CLng(0) '** UNCHANGED
        
        
        Call m_Form.GetMedOrderInfo(m_Form.txtRXOP_Number.Text, PhysicianName, PhysicianNCID, DaysOfSupply, Dose, Dosage_Form)
        m_Form.txtRXOP_PhysicianName.Text = PhysicianName
        m_Form.txtRXOP_OutsidePhysicianName.Enabled = PhysicianName = ".UNASSIGNED"
        
        m_Form.txtRXOP_PhysicianName.Tag = PhysicianNCID
        m_Form.txtRXOP_DaysOfSupply.Text = DaysOfSupply
        m_Form.txtRXOP_Dose.Text = Dose
        m_Form.txtRXOP_DosageForm.Text = Dosage_Form
        
        '** 7/6/05 - Frank -- Only allowing Pharmacist to enter new med order
        '** from med tab.  IV and IVPB are not allowed
        If CreateNewRx Then
            m_Form.AddNewRx
            m_Form.TabOne1.TabEnabled(3) = False
        End If
        
            
        '** Frank - 7/27/05
        If m_Form.lstOrders.Text = "BLANK LABEL" Then
            m_Form.txtRXOP_Number.Text = ""
            m_Form.txtRXOP_Medication.Text = ""
            m_Form.txtRxOP_ExpSig.Text = ""
            m_Form.txtRXOP_OutsidePhysicianName.Text = ""
            
            m_Form.txtRXOP_Number.Enabled = False
            m_Form.txtRXOP_Medication.Enabled = False
            m_Form.txtRxOP_ExpSig.Enabled = False
            m_Form.txtRXOP_OutsidePhysicianName.Enabled = False
        End If
        
          ' 125183 5/9/07 Quan
'        If (mUser.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Write) = False And _
'            mUser.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Read) = True) Then
'       Changed to requirement place the edit control on the Sign priv instead Write priv
        If (mUser.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Sign) = False And _
            mUser.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Read) = True) Then
            m_Form.txtRXOP_Dose.Enabled = False
            m_Form.txtRXOP_DosageForm.Enabled = False
            m_Form.txtRXOP_DaysOfSupply.Enabled = False
            m_Form.txtRXOP_Comments.Enabled = False
            m_Form.txtRXOP_SIG.Enabled = False
            m_Form.txtRXOP_Quantity.Enabled = False
            m_Form.txtRXOP_NumberRefills.Enabled = False
            m_Form.cmdDispense.Enabled = False
        End If
        
    End Select
    

    
    initLabels = True
    DoEvents
    
Func_Exit:
  On Error Resume Next
  mb_Loading = False
  Screen.MousePointer = vbDefault
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.initLabels", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug
End Function




Private Function GetOrderID(DrugName As String) As String

    '*** SCR 67259 - Frank
    '*** 3/31/05 Display Patient Order Type

    GetOrderID = ""
    
    If Not mo_Rst.EOF Then
      If Len(Trim$(DrugName)) > 0 Then  '** Skip BLANK LABEL
          mo_Rst.Find "ORDER_NAME = '" & DrugName & "'"
          If Not mo_Rst.EOF Then
               GetOrderID = mo_Rst.Fields("ORDER_ID")
          End If
      End If
    End If
End Function


' This module is used to retrieve values that have to be manipulated in some way. All other values are passed directly from the
' in memory values (recordsets, dictionaries etc)
' This module gets called if the XML dbfield value has a Pipe "|" delimeter in the value. ie CLINIC_NAME|0 or the value
' after the delimeter can be used to further define what to get ie CLINIC_NAME|CLINICNCID, CLINICNCID could be the value we are looking for or
' need to use to obtain a certain value

Private Function GetDefinedData(ByVal sDBFieldName As String, ByRef sCtrlVal As String) As Boolean
  Dim sTemp() As String
  
  On Error GoTo Func_Err
  
  GetDefinedData = False
  
  sTemp = Split(sDBFieldName, "|")
  
  If UBound(sTemp) > 0 Then
    
    Select Case UCase$(sTemp(0))
    Case "CLINIC_NAME"
      If Len(ms_ClinicName) > 0 Then
        sCtrlVal = ms_ClinicName
      End If
    Case "CLINIC_ADDRESS_TELEPHONE"
      If Len(ms_ClinicName) > 0 Then
        sCtrlVal = ms_ClinicName
      End If
      If Len(ms_ClinicAddress) > 0 Then
        sCtrlVal = sCtrlVal & " " & ms_ClinicAddress
      End If
      If Len(ms_ClinicPhone) > 0 Then
        sCtrlVal = sCtrlVal & " " & ms_ClinicPhone
      End If
    Case "CLINIC_ADDRESS"
      If Len(ms_ClinicName) > 0 Then
        sCtrlVal = ms_ClinicName
      End If
      If Len(ms_ClinicAddress) > 0 Then
        sCtrlVal = sCtrlVal & " " & ms_ClinicAddress
      End If
    Case "TELEPHONE"
      If Len(ms_ClinicPhone) > 0 Then
        sCtrlVal = ms_ClinicPhone
      End If
    Case "LAST_FIRST_MI"
      sCtrlVal = mo_PatientRec.Fields("Full_name").Value & vbNullString
    Case "FMP"
      sCtrlVal = mo_PatientRec.Fields("FMP").Value & vbNullString
    Case "AGE"
      sCtrlVal = mo_PatientRec.Fields("AGE").Value & vbNullString
    Case "DOB"
      sCtrlVal = Format$(mo_PatientRec.Fields("DOB").Value & vbNullString, ms_DateFormat)
    Case "GENDER"
      sCtrlVal = mo_PatientRec.Fields("SEX").Value & vbNullString
    Case "PATIENT_ID"
      sCtrlVal = left$(mo_PatientRec.Fields("Last_name").Value, 1) & vbNullString & right$(mo_PatientRec.Fields("SSN").Value, 4) & vbNullString
    Case "FMP_SSN"
      sCtrlVal = mo_PatientRec.Fields("FMP").Value & vbNullString & mo_PatientRec.Fields("SSN").Value & vbNullString
    Case "SSN_FMP"
      sCtrlVal = mo_PatientRec.Fields("SSN").Value & vbNullString & mo_PatientRec.Fields("FMP").Value & vbNullString
    Case "SSN"
      sCtrlVal = mo_PatientRec.Fields("SSN").Value & vbNullString
    Case "DOB"
      sCtrlVal = Format$(mo_PatientRec.Fields("DOB").Value & vbNullString, ms_DateFormat)
    Case "LS_BARCODE"
      sCtrlVal = mo_Rst.Fields(CStr(sTemp(1))).Value & vbNullString
    Case "RFF_BARCODE"
      Select Case CStr(sTemp(1))
      Case "FMP_SSN"
        sCtrlVal = mo_PatientRec.Fields("FMP").Value & vbNullString & mo_PatientRec.Fields("SSN").Value & vbNullString
      Case "SSN_FMP"
        sCtrlVal = mo_PatientRec.Fields("SSN").Value & vbNullString & mo_PatientRec.Fields("FMP").Value & vbNullString
      Case "SSN"
        sCtrlVal = mo_PatientRec.Fields("SSN").Value & vbNullString
      End Select
    Case "PRIORITY"
      sCtrlVal = GetPriValue(CLng(mo_Rst.Fields(CStr(sTemp(1))).Value))
    Case "BED"
      sCtrlVal = mo_BedInfoRst.Fields("ASSIGNED_BED_NUMBER").Value
    Case "BEDLOCATION"
      sCtrlVal = GetBedLocation(mo_BedInfoRst.Fields("WARD_NCID").Value)
    Case "BEDLOCATION_NUMBER"
      If Len(mo_BedInfoRst.Fields("WARD_NCID").Value) > 0 Then
        sCtrlVal = GetBedLocation(mo_BedInfoRst.Fields("WARD_NCID").Value) & "/Bed # " & mo_BedInfoRst.Fields("ASSIGNED_BED_NUMBER").Value
      End If
    Case Else
    End Select
    
    
  
    GetDefinedData = True
  
  End If

Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  Resume Func_Exit
  Resume 'Debug


End Function

Private Function LoadConfigXML() As Boolean
  
  On Error GoTo Func_Err
  
  LoadConfigXML = False
  
  If oDom Is Nothing Then
    Set oDom = New DOMDocument
  End If
  
  oDom.async = False
  oDom.validateOnParse = False
  oDom.resolveExternals = False
  
  oDom.Load gobjshared.RTDataPath & "\" & "CHCSII_LabelSetup.xml"

  If oDom.parseError Then
    Select Case oDom.parseError.errorCode
    'Missing XML File
    Case -2146697210
      MsgBox "Unable to locate the CHCSII Label Configuration XML file. Cannot continue without this file", vbCritical, "XML Label Config Error"
    'XML is not well formed
    Case Else
      MsgBox "XML file does not appear to be well formed." & vbCrLf & vbCrLf & "Reason: " & oDom.parseError.reason & "Line: " & oDom.parseError.linepos & vbCrLf & _
        "Suspect text: " & vbCrLf & "*******************" & vbCrLf & oDom.parseError.srcText & vbCrLf & "*******************", vbCritical, "XML Label Config Error"
    End Select
    Set oDom = Nothing
    GoTo Func_Exit
  End If
  
  LoadConfigXML = True

Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.LoadConfigXML", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug
End Function
Private Function PersistConfigXML() As Boolean
  Dim oNode As IXMLDOMElement
  On Error GoTo Func_Err
  
  PersistConfigXML = False
  
  If oDom Is Nothing Then
    GoTo Func_Exit
  End If
  
  Set oNode = oDom.documentElement
  
  If Not oNode Is Nothing Then
    oNode.setAttribute "LastTab", "Tab" & m_Form.TabOne1.CurrTab
  End If

  
  'Some how this is getting set to nothing mid stream
  If gobjshared Is Nothing Then
    Set gobjshared = New CWShared
  End If
  
  oDom.Save gobjshared.RTDataPath & "\" & "CHCSII_LabelSetup.xml"

  
  PersistConfigXML = True

Func_Exit:
  On Error Resume Next
  Set oNode = Nothing
  Exit Function
  
Func_Err:
  If Err.Number = -2147024891 Then
    If MsgBxARMd("Could not save the CHCSII Label values because the " & gobjshared.RTDataPath & "\" & _
      "CHCSII_LabelSetup.xml files is Write Protected." & vbCrLf & vbCrLf & _
      "Change the file to not be Write Protected by unchecking the 'Read Only' attribute of the file " & _
      "then press the 'Retry' button or 'Cancel' to exit without saving.", vbRetryCancel, "CHCSII Label Save Error") = vbRetry Then
      Err.Clear
      Resume
    Else
      Resume Func_Exit
    End If
  Else
    gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.PersistConfigXML", "CHCSII Label Error", vbExclamation
    Resume Func_Exit
    Resume 'Debug
  End If
End Function

' Module Description: Used to retrieve the Order Entry data depending on which tab is selected. Because OE_RAD, OE_LAB and OE_RX are similar
' in how they store data, one module level recordset can be used to retrieve data for all labels
Private Function GetListData() As Boolean
  Dim sSQL As String
  
  On Error GoTo Func_Err
  
  'Get the domain data based on the tab selected
  Select Case ml_Tab
  
  Case tab0
    sSQL = "SELECT * from OE_LAB where unit_number = " & ms_PatientUnitNumber
  Case Tab1
    sSQL = "SELECT * from OE_RAD where unit_number = " & ms_PatientUnitNumber
  Case tab2
    sSQL = "SELECT * from OE_RAD where unit_number = " & ms_PatientUnitNumber
  Case tab3
    sSQL = "SELECT OE_RX.*, M.FDB_MEDID from OE_RX LEFT JOIN MEDICATIONS M ON M.ORDERNUMBER = OE_RX.ORDER_ID WHERE OE_RX.unit_number = " & ms_PatientUnitNumber & " Order By ORDER_NAME ASC"
  Case tab4
    sSQL = "SELECT OE_RX.*, M.FDB_MEDID from OE_RX LEFT JOIN MEDICATIONS M ON M.ORDERNUMBER = OE_RX.ORDER_ID WHERE OE_RX.unit_number = " & ms_PatientUnitNumber & " Order By ORDER_NAME ASC"
  Case Tab5
    sSQL = ""
  Case Else
  End Select
  
  Set mo_Rst = goDBDAS.OpenRecordset(sSQL)

  'Return false and exit out if no data is found
  If Not mo_Rst Is Nothing Then
    If mo_Rst.BOF And mo_Rst.EOF Then
      GetListData = False
      GoTo Func_Exit
    End If
  Else
    GetListData = False
    GoTo Func_Exit
  End If

  GetListData = True

Func_Exit:
  On Error Resume Next
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.GetListData", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug
End Function

Private Sub LoadList()
  Dim lCnt As Long
  Dim selIndes As Long
  On Error GoTo Sub_Err
  
  'Clear the listbox
  m_Form.lstOrders.Clear
  
  mdic_ListIdxField.RemoveAll
  
  
  mo_Rst.MoveFirst
  
  lCnt = 0
  
  'Update listbox with database column defined in XML Config file
  'This gives the user the flexibility to change how the data is selected and filtered in the UI
    Do Until mo_Rst.EOF
  
        Select Case ml_Tab
            Case tab3, tab4
                '*** 4/19/05 - Frank
                '*** Include Inpatient Active Med Order also
                If Val(mo_Rst.Fields("PATIENT_TYPE").Value & vbNullString) = 1 Then
                    '*********************************
                    '** INPATIENT ACTIVE MED ORDER
'                     If (mo_Rst.Fields("STATUS_NCID").Value = NCID_PARTIAL) Then
                    If (mo_Rst.Fields("STATUS_NCID").Value = NCID_PARTIAL) Or (mo_Rst.Fields("STATUS_NCID").Value = NCID_ORDERSTATUS_ACTIVE) Then
                        mdic_ListIdxField.Add lCnt, mo_Rst.Fields("ORDER_ID").Value
                        m_Form.lstOrders.AddItem mo_Rst.Fields("ORDER_Name").Value & vbNullString
                        GoTo Continue_Next_Record
                    Else
                        GoTo Continue_Skip_Record
                    End If
                Else
                    '*********************************
                    '** OUTPATIENT PENDING MED ORDER
                
                    '*** 4/19/05 Frank
                    '*** For Multi-Med Order, Display only the PENDING MED ORDER (REFILL)
                    '*** Skip ACTIVE Med Order
                    If getMultiOrderCount(mo_Rst.Fields("ORDER_ID").Value) Then
                        
                        '** For Multiple Med Order, display only Pending Med Order
                        If mo_Rst.Fields("STATUS_NCID").Value = NCID_PARTIAL Then
                            mdic_ListIdxField.Add lCnt, mo_Rst.Fields("ORDER_ID").Value
                            m_Form.lstOrders.AddItem mo_Rst.Fields("ORDER_Name").Value & vbNullString
                            GoTo Continue_Next_Record
                        Else
                            GoTo Continue_Skip_Record
                        End If
                        
                    Else

                        If (mo_Rst.Fields("STATUS_NCID").Value = NCID_PARTIAL) Or _
                            (mo_Rst.Fields("STATUS_NCID").Value = NCID_ORDERSTATUS_ACTIVE) Or _
                            (mo_Rst.Fields("STATUS_NCID").Value = NCID_ORDERSTATUS_COMPLETE) Then
                        
                            m_Form.lstOrders.AddItem mo_Rst.Fields(mdic_ListDBField.Item(ml_Tab)).Value & vbNullString
                        
                            '<< Begin:SCR #47482;   Developer: Brian Mowbray 12/29/2003 02:08 PM
                            'This gives me a way to differentiate between duplicate entries
                            mdic_ListIdxField.Add lCnt, mo_Rst.Fields("ORDER_ID").Value
                        Else
                            GoTo Continue_Skip_Record
                        End If
                    End If
                End If
            Case Else
                '*** For Rad/Lab do not skip any records
        
                m_Form.lstOrders.AddItem mo_Rst.Fields(mdic_ListDBField.Item(ml_Tab)).Value & vbNullString
                
                '<< Begin:SCR #47482;   Developer: Brian Mowbray 12/29/2003 02:08 PM
                'This gives me a way to differentiate between duplicate entries
                mdic_ListIdxField.Add lCnt, mo_Rst.Fields("ORDER_ID").Value
            
        End Select
        

    
Continue_Next_Record:

        If LenB(mOrderId) > 0 And firstTime Then
            If mOrderId = mo_Rst.Fields("ORDER_ID").Value Then
                selIndes = lCnt
            End If
        End If
        
        lCnt = lCnt + 1

Continue_Skip_Record:
        mo_Rst.MoveNext
    
    Loop
  
    m_Form.lstOrders.AddItem "BLANK LABEL"
  
    If Not mCreateNewRx Then    '** When user click New Rx button from Order Mgt, don't select IV or IVPB med which result in display IV tab w/o New Rx feature implemented
        If LenB(mOrderId) > 0 And firstTime Then
            mb_Loading = False
            m_Form.lstOrders.ListIndex = selIndes
            SelectedIndex = selIndes  '--- SCR 72008  Sherry Wang  7/6/2005
            firstTime = False
        End If
    End If

Sub_Exit:
    On Error Resume Next
    Exit Sub
  
Sub_Err:
    gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.LoadOrdersList", "CHCSII Label Error", vbExclamation
    Resume Sub_Exit
    Resume 'Debug
End Sub


Public Sub LoadSelected()
  If Not mb_Loading Then
    initLabels False
  End If
End Sub


Public Function PrintLabel(oForm As Object) As Boolean
  Dim oCrystal As Crystal
  Dim msg As String
        
  Const CrystralReportMaxCharLimit As Integer = 254
  On Error GoTo Func_Err

  PrintLabel = False

  Set m_Form = oForm

  ml_Tab = m_Form.TabOne1.CurrTab
  
  DoEvents

  'SCR-134909 - FL - 7/17/07
  'Crystal Report 8.0/8.5 error, limit string to 254 to work around Crystal Report Crash on "string has too many characters."
  If ml_Tab = PharmacyOutPat Then
    If Len(m_Form.txtRxOP_ExpSig.Text) > CrystralReportMaxCharLimit Then
          msg = "The expanded SIG is more than " & CrystralReportMaxCharLimit & " characters, click Yes to truncate the Expanded SIG to " & CrystralReportMaxCharLimit & " characters, or click No to edit the order and shorten the SIG such that the Expanded SIG is less than " & CrystralReportMaxCharLimit & " characters."
          If MsgBox(msg, vbQuestion + vbYesNo, "Labels") = vbNo Then
              If m_Form.txtRXOP_SIG.Enabled Then
                  m_Form.txtRXOP_SIG.SetFocus
              End If
              Exit Function
          Else
              m_Form.txtRxOP_ExpSig.Text = Mid$(m_Form.txtRxOP_ExpSig.Text, 1, CrystralReportMaxCharLimit)
              m_Form.ExpSigCount.Caption = Len(m_Form.txtRxOP_ExpSig.Text)
          End If
    End If
  End If
  
''    SelectPrinter (oForm.cboPrinters.Text)
  LoadPrintSettings

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set mo_Report = Nothing
    If Not oCrystal.ReportOpen(mdic_Report.Item(ml_Tab), mo_Report) Then GoTo Func_Exit
    
    If Not oCrystal.AssignFormulas(moCollFormulas, mo_Report) Then GoTo Func_Exit
   
   ' SCR 120563 2/15/07 Quan
    'AdjustFormulaProperties mo_Report
    
    HandleApostrophe mo_Report

    '--- SCR 70955   Sherry Wang   5/18/2005
'    If Printer.DeviceName = "DYMO LabelWriter 330 Turbo-USB" Then
'        mo_Report.PaperOrientation = crLandscape
'        '--- SCR 88162  Sherry Wang   2/16/2006
'        '--- SCR 88350,  88352   Sherry Wang  2/19/2006
'        If mo_Report.PaperSize <> Printer.PaperSize Then
'            mo_Report.PaperSize = crPaperEnvelopeMonarch
'        End If
'    End If

    ' SCR 120558 2/15/07 Quan
    If SetLabelPaperSize = False Then
        Exit Function
    End If
    mo_Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port

    ' 132632, 133706 6/11/07 Quan
    If mo_Report.PrintingStatus.NumberOfPages > 1 Then
        msg = "This order will print across more than one medication label.  Please shorten the Expanded SIG or the Comments if one label is desired. Do you want to continue?"
        If m_Form.txtRXOP_SIG.Enabled = False Then
            msg = "This order will print across more than one medication label.  Please shorten the Expanded SIG or the Comments by editing the order in A/P if one label is desired.  Do you want to continue?"
        End If
        If MsgBox(msg, vbQuestion + vbYesNo, "Labels") = vbNo Then
            Set oCrystal = Nothing
            Exit Function
        End If
    End If
    
    If mi_PrintOption = PrintOnly Then
        If Not oCrystal.PrintReport(mo_Report) Then GoTo Func_Exit
    ElseIf mi_PrintOption = PrintPreview Then
        If Not oCrystal.DisplayReport(mo_Report) Then GoTo Func_Exit
    End If

PrintLabel = True

Func_Exit:
  Set oCrystal = Nothing
  Exit Function

Func_Err:
  MsgBox "Error in PrintLabel: " & Err.Description
  Resume Func_Exit
  Resume 'Debug
End Function
Private Function SetLabelPaperSize() As Boolean
' SCR 120558 2/15/07 Quan
'This sub determines if the correct label paper size exists for the selected printer.
'If so, the correct paper size is selected.

Dim PaperID As Integer
Dim sLabelSize As String

    sLabelSize = "30256 Shipping" '"30323 Shipping"
    SetLabelPaperSize = True
    PaperID = modPrinter.GetPaperID(sLabelSize) 'Look for Shipping Label

    If PaperID = -1 Then
        '--- SCR 135870  Sherry Wang  7/24/2007
        mo_Report.TopMargin = 300
        If MsgBox("The selected printer '" & Printer.DeviceName & "' does not support the '" & sLabelSize & "' label size." & vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion, "Continue printing?") = vbNo Then
            SetLabelPaperSize = False
        End If
    Else ' label was found, so set the settings.
         modPrinter.SetPrnt PaperID, vbPRORLandscape
         Printer.CurrentX = 0
         Printer.CurrentY = 0
'        Printer.Print "Printed the API way. This is a test. This is only a test. Please stand by."
'        Printer.EndDoc
    End If

End Function

'--- SCR 67595  Sherry Wang  3/12/2005
Private Sub HandleApostrophe(oReport As CRAXDRT.Report)
    Dim i As Integer
    Dim j As Integer

    On Error GoTo ErrHandler
    
    For i = 1 To oReport.Sections.Count
        For j = 1 To oReport.Sections.Item(i).ReportObjects.Count
            If left$(oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text, 1) = "'" Then
                oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text = Chr(34) & Mid(oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text, 2, Len(oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text) - 2) & Chr(34)
            ElseIf left$(oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text, 1) <> Chr(34) Then
                oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text = Chr(34) & oReport.Sections.Item(i).ReportObjects.Item(j).Field.Text & Chr(34)
            End If
        Next j
    Next i
    
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCHCSII_Labels.HandleApostrophe", App.EXEName, vbExclamation
Cleanup:

End Sub


Public Sub AdjustFormulaProperties(oReport As CRAXDRT.Report)
  Dim i As Integer
  Dim x As Integer
  On Error Resume Next

    '-- Set Formula Fields properties
    '-----------------------------------------------------------------------
    For i = 1 To oReport.Sections.Count
      For x = 1 To oReport.Sections(i).ReportObjects.Count
        If oReport.Sections(i).ReportObjects.Item(x).Kind = crFieldObject Then
            If UCase$(left$(oReport.Sections(i).ReportObjects.Item(x).Field.FormulaFieldName, 10)) = "LINEHEADER" Then
              oReport.Sections(i).ReportObjects.Item(x).Font.Size = ml_HeaderFontSize
              ' SCR 98099 Quan 6/22/06
            ElseIf UCase$(left$(oReport.Sections(i).ReportObjects.Item(x).Field.FormulaFieldName, 5)) = "LINE7" Then
              oReport.Sections(i).ReportObjects.Item(x).Font.Size = ml_DetailFontSize - 2
              
            ElseIf UCase$(left$(oReport.Sections(i).ReportObjects.Item(x).Field.FormulaFieldName, 4)) = "LINE" Then
              oReport.Sections(i).ReportObjects.Item(x).Font.Size = ml_DetailFontSize
            End If
                        
        End If
      Next
    Next
End Sub


'This is called once when the class loads to retrieve the clinic address info
Private Sub GetClinicInfo()
  Dim sClinicNCID As String
  Dim sSQL As String
  On Error GoTo Sub_Err
  
  
  sSQL = "select " & _
            "o.APVALUE " & _
        "from " & _
            "OSH_INI_SETTINGS o " & _
        "where " & _
            "o.APKEY = '" & "CLINIC_NCID" & "'"
  
   goDBSqlOps.Execute sSQL
  
  If goDBSqlOps.EOF Then
    ms_ClinicName = vbNullString
    ms_ClinicAddress = vbNullString
    ms_ClinicPhone = vbNullString
    GoTo Sub_Exit
  Else
    sClinicNCID = goDBSqlOps.Value("APVALUE")
  End If
  
  
  sSQL = "select " & _
            "c.Name, c.STREET_ADDRESS, c.BUILDING_NUMBER, c.CITY, c.STATE, c.ZIP, c.TELEPHONE_NUMBER " & _
        "from " & _
            "Clinic c " & _
        "where " & _
            "c.ncid = " & sClinicNCID
  
   goDBSqlOps.Execute sSQL
  
  If goDBSqlOps.EOF Then
    ms_ClinicName = vbNullString
    ms_ClinicAddress = vbNullString
    ms_ClinicPhone = vbNullString
  Else
    ms_ClinicName = goDBSqlOps.Value("Name")
    ms_ClinicAddress = goDBSqlOps.Value("STREET_ADDRESS") & " " & goDBSqlOps.Value("BUILDING_NUMBER") & " " & _
                        goDBSqlOps.Value("CITY") & ", " & goDBSqlOps.Value("STATE") & " " & goDBSqlOps.Value("ZIP")
    ms_ClinicPhone = goDBSqlOps.Value("TELEPHONE_NUMBER")
  End If

Sub_Exit:
  On Error Resume Next
  Exit Sub
  
Sub_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.GetClinicInfo", "CHCSII Label Error", vbExclamation
  Resume Sub_Exit
  Resume 'Debug
End Sub

Public Property Get ReportObject() As CRAXDRT.Report
  Set ReportObject = mo_Report
End Property

Public Property Set ReportObject(ByVal vValueIn As CRAXDRT.Report)
  Set mo_Report = vValueIn
End Property

Private Sub LoadPatientData()

  On Error GoTo Sub_Err
  
  If mo_PatientDemographics Is Nothing Then
    Set mo_PatientDemographics = New PatientDemographics
  End If
  
  If Not mo_PatientRec Is Nothing Then
    mo_PatientRec.MoveFirst
    If mo_PatientRec.Fields("Unit_Number").Value <> ms_PatientUnitNumber Then
      If Not mo_PatientDemographics.RecordGetDemographics(ms_PatientUnitNumber, mo_PatientRec) Then GoTo Sub_Exit
    End If
  Else
    If Not mo_PatientDemographics.RecordGetDemographics(ms_PatientUnitNumber, mo_PatientRec) Then GoTo Sub_Exit
  End If
  
  With m_Form
    Select Case ml_Tab
    Case tab0
      .txtLS_Name.Text = mo_PatientRec.Fields("Full_name").Value & vbNullString
      .txtLS_Name.Tag = IIf(Len(.txtLS_Name.Text) > 0, "LOCK", "")
      .txtLS_FMP.Text = mo_PatientRec.Fields("FMP").Value & vbNullString
      .txtLS_FMP.Tag = IIf(Len(.txtLS_FMP.Text) > 0, "LOCK", "")
      .txtLS_Gender.Text = mo_PatientRec.Fields("Sex").Value & vbNullString
      .txtLS_Gender.Tag = IIf(Len(.txtLS_Gender.Text) > 0, "LOCK", "")
      .txtLS_Age.Text = mo_PatientRec.Fields("AGE").Value
      .txtLS_Age.Tag = IIf(Len(.txtLS_Age.Text) > 0, "LOCK", "")
'      .txtLS_SSN.Text = mo_PatientRec.Fields("SSN").Value
'      .txtLS_SSN.Tag = IIf(Len(.txtLS_SSN.Text) > 0, "LOCK", "")
'      If Len(mo_PatientRec.Fields("DOB").Value & vbNullString) > 0 Then
'        .txtLS_DOB.Text = Format$(mo_PatientRec.Fields("DOB").Value, ms_DateFormat)
'      End If
'      .txtLS_DOB.Tag = IIf(Len(.txtLS_DOB.Text) > 0, "LOCK", "")
    Case Tab1
      .txtRFF_Name.Text = mo_PatientRec.Fields("Full_name").Value & vbNullString
      .txtRFF_Name.Tag = IIf(Len(.txtRFF_Name.Text) > 0, "LOCK", "")
      .txtRFF_FMP_SSN.Text = mo_PatientRec.Fields("FMP").Value & vbNullString & mo_PatientRec.Fields("SSN").Value & vbNullString
      .txtRFF_FMP_SSN.Tag = IIf(Len(.txtRFF_FMP_SSN.Text) > 0, "LOCK", "")
      .lblRFF_Barcode.Caption = mo_PatientRec.Fields("FMP").Value & vbNullString & mo_PatientRec.Fields("SSN").Value & vbNullString
    Case tab2
      .txtRFL_Name.Text = mo_PatientRec.Fields("Full_name").Value & vbNullString
      .txtRFL_Name.Tag = IIf(Len(.txtRFL_Name.Text) > 0, "LOCK", "")
      .txtRFL_FMP_SSN.Text = mo_PatientRec.Fields("FMP").Value & vbNullString & mo_PatientRec.Fields("SSN").Value & vbNullString
      .txtRFL_FMP_SSN.Tag = IIf(Len(.txtRFL_FMP_SSN.Text) > 0, "LOCK", "")
      .txtRFL_Gender.Text = mo_PatientRec.Fields("Sex").Value & vbNullString
      .txtRFL_Gender.Tag = IIf(Len(.txtRFL_Gender.Text) > 0, "LOCK", "")
      .txtRFL_Age.Text = mo_PatientRec.Fields("AGE").Value
      .txtRFL_Age.Tag = IIf(Len(.txtRFL_Age.Text) > 0, "LOCK", "")
    Case tab3
      .txtRXIV_Name.Text = mo_PatientRec.Fields("Full_name").Value & vbNullString
      .txtRXIV_Name.Tag = IIf(Len(.txtRXIV_Name.Text) > 0, "LOCK", "")
    Case tab4
      .txtRXOP_Name.Text = mo_PatientRec.Fields("Full_name").Value & vbNullString
'      .txtRXOP_Name.Tag = IIf(Len(.txtRXOP_Name.Text) > 0, "LOCK", "")
    Case Tab5
    
    End Select
  End With

Sub_Exit:
  On Error Resume Next
  Exit Sub
  
Sub_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.LoadPatientData", "CHCSII Label Error", vbExclamation
  Resume Sub_Exit
  Resume 'Debug

End Sub

Private Sub SaveLabelData()
  Dim oNode As IXMLDOMElement
  Dim oNodeElement As IXMLDOMElement
  Dim sAttribute As String
  Dim sValue As String
  On Error GoTo Sub_Err

  If oDom Is Nothing Then
    GoTo Sub_Exit
  End If
  
  'Load Label Data from XML for Tab selected
  Set oNode = oDom.selectSingleNode(("//CHCSII_Labels/Tab" & CStr(ml_LastTab)) & "/Controls")
  
  If Not oNode Is Nothing Then
      'Run through the xml tree to set the controls with the cached data and/or attributes
      For Each oNodeElement In oNode.childNodes
        'Save control data
        sAttribute = oNodeElement.getAttribute("Name")
        If m_Form(sAttribute).Tag <> "LOCK" And left$(sAttribute, 3) = "txt" Then
          sValue = m_Form(sAttribute)
          If Len(sValue) > 0 Then
            oNodeElement.setAttribute "Value", m_Form(sAttribute) & vbNullString
          End If
        End If
      Next
  End If
  

Sub_Exit:
  On Error Resume Next
  Set oNode = Nothing
  Set oNodeElement = Nothing
  Exit Sub
  
Sub_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.SaveLabelData", "CHCSII Label Error", vbExclamation
  Resume Sub_Exit
  Resume 'Debug
End Sub

Private Sub SetCtrlPermission()
  On Error Resume Next
  Dim oCtrl As Control
  
  For Each oCtrl In m_Form.Controls
    'Set the controls based on the tag property
    oCtrl.Enabled = IIf(oCtrl.Tag = "LOCK", False, True)
    oCtrl.Locked = IIf(oCtrl.Tag = "LOCK", True, False)
    
    'Set the controls based on dictionary
    If mdic_CtrlLock.Exists(oCtrl.Name) Then
      oCtrl.Enabled = IIf(mdic_CtrlLock.Item(oCtrl.Name) = "True", False, True)
      oCtrl.Locked = IIf(mdic_CtrlLock.Item(oCtrl.Name) = "True", True, False)
    End If
  Next
End Sub

Private Sub LoadControls()
  Dim oNode As IXMLDOMElement
  Dim oNodeElement As IXMLDOMElement
  Dim sCtrlLockVal As String
  On Error GoTo Sub_Err
  
  If oDom Is Nothing Then
    GoTo Sub_Exit
  End If
  
  'Need to reset the Control Dictionary Each time this is loaded
  Set mdic_CtrlLock = Nothing
  Set mdic_CtrlLock = New Dictionary
  
  'Load Label Data from XML for Tab selected
  Set oNode = oDom.selectSingleNode(("//CHCSII_Labels/Tab" & CStr(ml_Tab)) & "/Controls")
  If Not mo_Rst Is Nothing Then
    '<< Begin:SCR #47482;   Developer: Brian Mowbray 12/29/2003 02:08 PM
    'Filter the data based on the field used to populate the listbox

    If mdic_ListIdxField.Exists(m_Form.lstOrders.ListIndex) Then
      mo_Rst.Filter = "ORDER_ID='" & mdic_ListIdxField.Item(m_Form.lstOrders.ListIndex) & "'"
    Else
      mo_Rst.Filter = mdic_ListDBField.Item(ml_Tab) & "='" & m_Form.lstOrders.Text & "'"
    End If

  '>> End: SCR #47482;
  
    '-- SCR 188798 (B1 meds have unlimited refills in B2)
    '-- Set the FDB_Med_Id global variable (used to determine if Print, Fill, Dispense buttons should be enabled
    On Error Resume Next
    If mo_Rst.RecordCount > 0 Then
        gFDB_Med_Id = mo_Rst.Fields("FDB_MEDID") & vbNullString
    Else
        gFDB_Med_Id = " "
    End If
    On Error GoTo Sub_Err
  End If
  
  
  
  If Not oNode Is Nothing Then
      'Run through the xml tree to set the controls with the cached data and/or attributes
      For Each oNodeElement In oNode.childNodes
        'Update control
        UpdateControl oNodeElement.getAttribute("Name"), oNodeElement.getAttribute("Value"), oNodeElement.getAttribute("dbfield"), CLng(NullToZero(oNodeElement.getAttribute("MaxLen")))
        'Add control specific locking information to the dictionary for later use.
        sCtrlLockVal = oNodeElement.getAttribute("Lock") & vbNullString
        If Len(sCtrlLockVal) > 0 Then
          mdic_CtrlLock.Add oNodeElement.getAttribute("Name"), sCtrlLockVal
        End If
      Next
  End If

Sub_Exit:
  On Error Resume Next
  Set oNode = Nothing
  Set oNodeElement = Nothing
  Exit Sub
  
Sub_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.LoadFromXML", "CHCSII Label Error", vbExclamation
  Resume Sub_Exit
  Resume 'Debug
End Sub
Private Function NullToZero(ByVal vVal As Variant) As String
  On Error Resume Next
  
  If IsNull(vVal) Then
    NullToZero = "0"
  ElseIf Len(Trim$(vVal)) = 0 Then
    NullToZero = "0"
  Else
    NullToZero = vVal
  End If
  
End Function

Private Sub LoadPrintSettings()
  Dim oNode As IXMLDOMElement
  Dim oNodeElement As IXMLDOMElement
  Dim oItemElement As IXMLDOMElement
  Dim sLine As String
  Dim sLineHeader As String
  Dim lHeaderSpaceL As Long
  Dim lHeaderSpaceR As Long
  Dim lHeaderLength As Long
  Dim sText As String
  Dim i As Integer
  
  Dim sSSN As String
  Dim iSSNBegin As Integer

  On Error GoTo Sub_Err
  
  If oDom Is Nothing Then
    GoTo Sub_Exit
  End If
  
'-- Set Formulas
'-----------------------------------------------------------------------------
  Set moCollFormulas = Nothing
  Set moCollFormulas = New Collection
        
  Const LABEL_LINE_LEN As Integer = 60
  
  'Load Label Data from XML for Tab selected
  Set oNode = oDom.selectSingleNode(("//CHCSII_Labels/Tab" & CStr(ml_Tab)) & "/LabelPrintSetting")
  
  If Not oNode Is Nothing Then
      'Run through the xml tree to set the controls with the cached data and/or attributes
      For Each oNodeElement In oNode.childNodes
        sLine = vbNullString
        sLineHeader = vbNullString
        i = 0
        For Each oItemElement In oNodeElement.childNodes
          i = i + 1
          'format the data time control values with the default data format
          If UCase$(left$(oItemElement.Text, 3)) = "DTP" Then
            '--- SCR 135440  Sherry Wang   7/17/2007
            If InStr(UCase(oItemElement.Text), "TIME") > 0 Then
                sText = Format$(m_Form(oItemElement.Text), ms_DateTimeFormat)
            Else
                sText = Format$(m_Form(oItemElement.Text), ms_DateFormat)
            End If
          Else
            sText = m_Form(oItemElement.Text)
          End If

          ' SCR 120564 2/15/07 Quan
          lHeaderSpaceL = CLng(oItemElement.getAttribute("leftmargin"))
          lHeaderSpaceR = CLng(oItemElement.getAttribute("rightmargin"))
          
          If lHeaderSpaceL < 0 Then lHeaderSpaceL = 0
          If lHeaderSpaceR < 0 Then lHeaderSpaceR = 0
          
          'Add field header and formating information from the xml file
          sLineHeader = sLineHeader & Space(lHeaderSpaceL) & Trim(oItemElement.getAttribute("Header")) & Space(lHeaderSpaceR)
          If Len(sLineHeader) < LABEL_LINE_LEN And i = oNodeElement.childNodes.Length Then
            sLineHeader = sLineHeader + Space(LABEL_LINE_LEN - Len(sLineHeader))
          ElseIf Len(sLineHeader) >= LABEL_LINE_LEN Then
            sLineHeader = left$(sLineHeader, LABEL_LINE_LEN)
          End If
                           
          If Len(oItemElement.Text) > 0 Then
            'Add field formating information from the xml file
            sLine = sLine & Space(lHeaderSpaceL) & Trim(sText)
            If Len(sLineHeader) - Len(sLine) > 0 Then
                sLine = sLine & Space(Len(sLineHeader) - Len(sLine))
            Else
                If oNodeElement.childNodes.Length > 1 Then
                    sLine = left$(sLine, Len(sLineHeader))
                End If
            End If
  
          End If
        Next
        If Len(sLineHeader) > 0 Then
          moCollFormulas.Add "'" & sLineHeader & "'", Trim("LineHeader" & oNodeElement.getAttribute("pos"))
        End If
        If Len(sLine) > 0 Then
          '--- SCR 67595   Sherry Wang  3/13/2005  Handle double quotes in the string
          sLine = Replace(sLine, Chr(34), Chr(34) & Chr(34))
          moCollFormulas.Add "'" & sLine & "'", Trim(oNodeElement.nodeName & oNodeElement.getAttribute("pos"))
        End If
      Next
  End If

Sub_Exit:
  On Error Resume Next
  Set oNode = Nothing
  Set oNodeElement = Nothing
  Set oItemElement = Nothing
  Exit Sub
  
Sub_Err:
  If Err.Number <> 728 Then
    'Add the error to the label
    sLine = sLine & Err.Description
  ElseIf Err.Number = 730 Then
    sText = oItemElement.Text
  End If
  Err.Clear
  Resume Next 'Debug
End Sub



Private Sub UpdateControl(sCtrlLabel As String, sCtrlVal As String, sDBFieldName As String, lMaxLen As Long)
  Dim oCtrl As Control
  Dim sDB_Val As String
  Dim sTemp As String
  On Error Resume Next

  For Each oCtrl In m_Form.Controls
    If oCtrl.Name = sCtrlLabel Then
      If Len(sDBFieldName) <= 0 Then
        If left(sCtrlVal, 1) = "T" Then
          If Len(sCtrlVal) > 1 Then
            If Mid$(sCtrlVal, 2, 1) = "-" Then
              sCtrlVal = Format$(Date - CLng(Mid(sCtrlVal, 3)), ms_DateFormat)
            Else
              sCtrlVal = Format$(Date + CLng(Mid(sCtrlVal, 3)), ms_DateFormat)
            End If
          Else
            sCtrlVal = Format$(Now, ms_DateFormat)
          End If
        End If
      Else
        If GetDefinedData(sDBFieldName, sCtrlVal) Then
          oCtrl.Tag = IIf(Len(sCtrlVal) > 0, "LOCK", "")
        Else
          sDB_Val = mo_Rst.Fields(sDBFieldName).Value & vbNullString
          If Len(sDB_Val) > 0 Then
            sCtrlVal = sDB_Val
          End If
          oCtrl.Tag = IIf(Len(sDB_Val) > 0, "LOCK", "")
        End If
        oCtrl = sCtrlVal
      End If
      If oCtrl.Tag <> "LOCK" Then
        oCtrl = sCtrlVal
      End If
      If UCase$(left$(oCtrl.Name, 3)) = "TXT" Then
        
        
        If lMaxLen > 0 Then
          oCtrl.MaxLength = lMaxLen
        Else
            Select Case UCase$(oCtrl.Name)
                '** SCR 85951 - 2/2/06 - Frank
                '** Limit Phone textbox to max 40 chars.
                Case UCase$("txtRXIV_MTFPhone"), UCase$("txtRXOP_MTFPhone")
                  oCtrl.MaxLength = 40
                Case UCase$("txtRxOP_ExpSig")
                  ' SCR 133402, jq, 6/22/07
                  ' don't mess with this... let it remain at 0 (unlim)
                Case Else
                  oCtrl.MaxLength = 80
            End Select
        End If
        
        ' SCR 120563 2/15/07 Quan
        If oCtrl.Name = "txtRXOP_SIG" Then
             oCtrl.MaxLength = 120
        End If
           
      End If
    End If
  Next


Sub_Exit:
  On Error Resume Next
  Set oCtrl = Nothing
  Exit Sub
  
  
Sub_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.UpdateControl", "CHCSII Label Error", vbExclamation
  Resume Sub_Exit
  Resume 'Debug

End Sub

Private Function SetupTabLabels() As Boolean
  Dim oNode As IXMLDOMElement
  Dim oNodeElement As IXMLDOMElement
  Dim sTabs As String
  Dim sLastChild As String
  Dim lCnt As Long
  On Error GoTo Func_Err

  SetupTabLabels = False

  If oDom Is Nothing Then
    GoTo Func_Exit
  End If
  
  lCnt = 0
  
  'This sets the node to the root node of the document
  Set oNode = oDom.documentElement
  
  'Get default values from the XML config file
  ms_FormCaption = oNode.getAttribute("FormCaption")
  ms_DateFormat = oNode.getAttribute("DateFormat")
  '--- SCR 135440  Sherry Wang   7/17/2007
  ms_DateTimeFormat = oNode.getAttribute("DateTimeFormat")
  ml_HeaderFontSize = CLng(oNode.getAttribute("HeaderFontSize"))
  ml_DetailFontSize = CLng(oNode.getAttribute("DetailFontSize"))
  
  
  'Sets the tabs names
  If Not oNode Is Nothing Then
      sLastChild = oNode.lastChild.nodeName
      For Each oNodeElement In oNode.childNodes
        If sLastChild <> oNodeElement.nodeName Then
          sTabs = sTabs & oNodeElement.getAttribute("Tabname") & "|"
        Else
          sTabs = sTabs & oNodeElement.getAttribute("Tabname")
        End If
        'Loads disctionaries with data that will be used later
        mdic_Report.Add lCnt, oNodeElement.getAttribute("ReportName")
        mdic_ListDBField.Add lCnt, oNodeElement.getAttribute("ListDBField")
        lCnt = lCnt + 1
      Next
      m_Form.TabOne1.Caption = sTabs
  End If
    
  If mOrderType = "LAB" Then
    m_Form.TabOne1.CurrTab = LabSpecimen
  ElseIf mOrderType = "RAD" Then
    m_Form.TabOne1.CurrTab = RadFilm
  ElseIf mOrderType = "MED" Then
    If LenB(gobjshared.CmdLineSwitches("INPATIENT2")) > 0 Then
'        m_Form.UserNotClickTab = True
        m_Form.TabOne1.CurrTab = PharmacyOutPat
    Else
        m_Form.TabOne1.CurrTab = PharmacyIV
    End If
  ElseIf ml_StartingTab = LastKnown Then
    m_Form.TabOne1.CurrTab = CLng(right(oNode.getAttribute("LastTab"), 1))
  Else
    m_Form.TabOne1.CurrTab = ml_StartingTab
  End If
  
  ' Quan SCR 78764
  If mCreateNewRx Then
    If LenB(gobjshared.CmdLineSwitches("INPATIENT2")) > 0 Then
'        m_Form.UserNotClickTab = True
        m_Form.TabOne1.CurrTab = PharmacyOutPat
    Else
        m_Form.TabOne1.CurrTab = PharmacyIV
    End If
  End If
  
  SetupTabLabels = True

Func_Exit:
  On Error Resume Next
  Set oNode = Nothing
  Set oNodeElement = Nothing
  Exit Function
  
Func_Err:
  gobjshared.ShowVBError Err, Err.Description, "CHCSII_Forms.FormCHCSII_Labels.SetupTabLabels", "CHCSII Label Error", vbExclamation
  Resume Func_Exit
  Resume 'Debug
End Function



Private Sub Class_Initialize()
    Dim sReturn As String
    mb_OrderManager = False 'SCR 131474 - LCI - 06/08/2007
    
    cwobjCWComm.Message cwiFOREGROUND_APPLICATION, sReturn, NCID_CW_APPLICATION, NCID_GENERAL_REPORTS_APPLICATION
    
    If sReturn = NCID_ORDER_MANAGER_INPATIENT_APPLICATION Then
        Dim OrderMgt As Object
        Set OrderMgt = cwobjCWComm.InitializeOLEServer(NCID_ORDER_MANAGER_INPATIENT_APPLICATION)
        'SCR 131474 - LCI - 06/08/2007: Used to determine if user is using Order Manager module
        mb_OrderManager = True
        If Not OrderMgt Is Nothing Then
            mOrderId = OrderMgt.OrderId
            mOrderType = OrderMgt.OrderType
            firstTime = True
        End If
        Set OrderMgt = Nothing
    End If
    
  If gobjshared Is Nothing Then
    Set gobjshared = New CWShared
  End If
  
  '*** 7/21/05 - Frank
  If goDBDAS Is Nothing Then Set goDBDAS = goDBConn.CHCSII_DAS(Auto)
  
  If mo_Rst Is Nothing Then
    Set mo_Rst = New ADODB.Recordset
  End If
  
  If mo_BedInfoRst Is Nothing Then
    Set mo_BedInfoRst = New ADODB.Recordset
  End If
  
  Set mo_PatientDemographics = New PatientDemographics
  
  Set mdic_Report = New Dictionary
  
  Set mdic_DBField = New Dictionary
  
  Set mdic_ListDBField = New Dictionary
  Set mdic_ListIdxField = New Dictionary
  
  Set mdic_CtrlLock = New Dictionary
  
  ml_StartingTab = LastKnown
  
  LoadConfigXML
  
  GetClinicInfo
  
End Sub


Private Sub Class_Terminate()
  On Error Resume Next
  
  ml_LastTab = ml_Tab
  
  SaveLabelData
  
  PersistConfigXML
  
  Set oDom = Nothing
  Set mo_PatientDemographics = Nothing
  Set mdic_Report = Nothing
  Set mdic_DBField = Nothing
  Set mdic_ListDBField = Nothing
  Set mdic_ListIdxField = Nothing
  Set mdic_CtrlLock = Nothing
  Set mo_PatientRec = Nothing
  Set mo_Rst = Nothing
  Set mo_BedInfoRst = Nothing
  Set mo_Crystal = Nothing
  Set moCollRecSet = Nothing
  Set moCollFormulas = Nothing
  
End Sub



Public Property Get XMLDOM() As DOMDocument
  Set XMLDOM = oDom
End Property

Public Property Set XMLDOM(ByVal vValueIn As DOMDocument)
  Set oDom = vValueIn
End Property

Public Property Get TabLabels() As String
  TabLabels = ms_TabLabels
End Property

Public Property Let TabLabels(ByVal vValueIn As String)
  ms_TabLabels = vValueIn
End Property

Public Property Get IsLoading() As Boolean
  IsLoading = mb_Loading
End Property

Public Property Let IsLoading(ByVal vValueIn As Boolean)
  mb_Loading = vValueIn
End Property

Public Property Get PatientUnitNumber() As String
  PatientUnitNumber = ms_PatientUnitNumber
End Property

Public Property Let PatientUnitNumber(ByVal vValueIn As String)
  ms_PatientUnitNumber = vValueIn
End Property

Public Property Get PrintOption() As Integer
  PrintOption = mi_PrintOption
End Property

Public Property Let PrintOption(ByVal vValueIn As Integer)
  mi_PrintOption = vValueIn
End Property



Public Sub SetStartingTab(ByVal iTab As LabelTabs)
  ml_StartingTab = iTab
End Sub

Public Sub SetStartingListValue(ByVal sListVal As String)
  On Error Resume Next
  If Len(sListVal) > 0 Then
    m_Form.lstOrders.Text = sListVal
  Else
    If Not mb_Loading Then
      If m_Form.lstOrders.ListCount > 0 Then
        m_Form.lstOrders.ListIndex = 0
      End If
    End If
  End If
End Sub


Private Function GetPatientType(OrderId As String) As String

    '*** SCR 67259 - Frank
    '*** 3/31/05 Display Patient Order Type
    
    On Error GoTo Err_Handler
    
    GetPatientType = ""
       
    If Not mo_Rst.EOF Then
      
      '** SCR   67309  - Frank - 4/15/05
      If (Len(OrderId) > 0) Then   '** Skip BLANK LABEL
          mo_Rst.Find "ORDER_ID = '" & OrderId & "'"
          If Not mo_Rst.EOF Then
             If "" & mo_Rst.Fields("PATIENT_TYPE") = "1" _
               Or UCase$("" & mo_Rst.Fields("PATIENT_TYPE")) = "I" Then
               GetPatientType = "INPATIENT ORDER"
''               GetPatientType = "Inpatient Order"
             Else
               GetPatientType = "OUTPATIENT ORDER"
''               GetPatientType = "Outpatient Order"
             End If
          End If
      End If
    End If
    
    Exit Function
Err_Handler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "FormCHCSII_Labels.GetPatientType", "Label Printing", vbExclamation)
    
End Function




'--- SCR 70956   Sherry Wang  5/18/2005
'--- Calculate the space between two headers
'--- The font size for header is 4, but the font size for content is 7
'--- In order to decide how many spaces we need between two headers in the same row
'--- based on the first column of the context, we need to calculate by number, spaces
'--- and characters.
Private Function CalculateSpaces(ByVal sText As String) As Long
    Dim lLen As Long
    Dim i As Integer
    Dim sChar As String
    
    On Error GoTo ErrHandler
    
    lLen = 0
    
    For i = 1 To Len(Trim$(sText))
        sChar = Mid(Trim$(sText), i, 1)
        If sChar = "0" Or sChar = "1" Or sChar = "2" _
           Or sChar = "3" Or sChar = "4" Or sChar = "5" _
           Or sChar = "6" Or sChar = "7" Or sChar = "8" Or sChar = "9" Then
            lLen = lLen + 1.75
        ElseIf sChar = " " Or sChar = "," Then
            lLen = lLen + 1
        ElseIf UCase(sChar) = sChar Then  '--- Upper Case
            lLen = lLen + 4
        Else
            lLen = lLen + 2
        End If
        
    Next i
    
    CalculateSpaces = Round(lLen)
    
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormCHCSII_Labels.CalculateSpaces", "sharedGlobal", vbExclamation
        
End Function


