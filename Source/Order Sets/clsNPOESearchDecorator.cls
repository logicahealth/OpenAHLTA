VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPOESearchDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "clsNPOESearchDecorator"

Private Const CAPTION_VERTICAL_OFFSET = 60
Private Const COMBO_HORIZONTAL_OFFSET = 225
Private Const BUTTON_HORIZONTAL_OFFSET = 45

'GUI Controls being decorated by this class
Private WithEvents mlblCaption As Label
Attribute mlblCaption.VB_VarHelpID = -1
Private WithEvents mcmdSearch As CommandButton
Attribute mcmdSearch.VB_VarHelpID = -1
Private WithEvents mcboProviders As ComboBox
Attribute mcboProviders.VB_VarHelpID = -1

'Internal variables
Private mobjProviderSearch As CHCSII_NonProvider.clsNonProvider 'Provides non-providers the
Private mrsSearchedProviders As ADODB.Recordset                 'Contains the list of providers
Private mbInit As Boolean                                       'inidicates if init has been called
Private msSelectedProviderIEN As String                         'IEN of the currently selected provider
Private mbVisible As Boolean                                    'Indicates whether the NPOE selection visible
Private mbEnabled As Boolean                                    'Indicates whether the NPOE selection is enabled
Private Sub Class_Terminate()
   
   'Destroy the search recordset
   If Not mrsSearchedProviders Is Nothing Then
      If mrsSearchedProviders.State = adStateOpen Then
         mrsSearchedProviders.Close
      End If
      Set mrsSearchedProviders = Nothing
   End If
   
   'Destroy the NPOE Search control
   Set mobjProviderSearch = Nothing
   
   'Destroy the controls
   Set mlblCaption = Nothing
   Set mcmdSearch = Nothing
   Set mcboProviders = Nothing

End Sub

Public Sub Init(ByRef lblCaption As Label, ByRef cboProviderCombo As ComboBox, ByRef cmdSearchButton As CommandButton)

On Error GoTo Error_Handler

   Set mlblCaption = lblCaption
   Set mcboProviders = cboProviderCombo
   Set mcmdSearch = cmdSearchButton
   mbInit = True
   
   ShowControls
   
   FillProvidersList
   
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".Init", Exe, vbExclamation)
   
End Sub

Private Sub InitNPOE()

On Error GoTo Error_Handler

   Set mobjProviderSearch = New CHCSII_NonProvider.clsNonProvider
   Set mrsSearchedProviders = New ADODB.Recordset
                   
   If gobjShared.IsAppMode(modeCHCSI_GUI) Then
       mobjProviderSearch.ConfigObject = cwobjConfig
   End If
   
   mobjProviderSearch.OpenNonProvider cwobjLogonEnviron, cwobjCWComm, gobjUser
   
   'Check to see if the provider selection controls should be displayed
   If mobjProviderSearch.SignableProvider Or gobjShared.IsAppMode(modeTheater) Then
      'JT SCR 42844 - NPOE selection is disabled for Theater
      DisableControls
   Else
      'Show the controls and retreive provider list
      ShowControls
   End If

   'Get the provider list
   If gobjShared.IsAppMode(modeTheater) Then
      'JT SCR 42844 - NPOE selection is disabled for Theater
      mcboProviders.AddItem gobjUser.UserName
      mcboProviders.ListIndex = 0
   Else
      mobjProviderSearch.getProviders mrsSearchedProviders
   End If
   
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".InitNPOE", Exe, vbExclamation)
   
End Sub

Public Sub FillProvidersList()
       
On Error GoTo Error_Handler

   If mrsSearchedProviders Is Nothing Then
      InitNPOE
   End If
              
   If mrsSearchedProviders.State <> adStateClosed Then
      
      'JT SCR 42844 - Only Clear when provider list is not initialized otherwise the control with have the user name
      mcboProviders.Clear
      msSelectedProviderIEN = ""
      
      If Not mrsSearchedProviders.EOF And Not mrsSearchedProviders.BOF Then
         
         mrsSearchedProviders.MoveFirst
         Do Until mrsSearchedProviders.EOF
             'Store the Name and the FLD_ORDER into the Combo...
             If mrsSearchedProviders.Fields("SIGNATURE_CLASS").Value > 1 Then
                 'mcboProviders.AddItem mrsSearchedProviders.Fields("NAME").Value & "^" & mrsSearchedProviders.Fields("IEN").Value
                 mcboProviders.AddItem mrsSearchedProviders.Fields("NAME").Value
                 mcboProviders.ItemData(mcboProviders.NewIndex) = mrsSearchedProviders.Fields("FLD_ORDER").Value
             End If
             mrsSearchedProviders.MoveNext
         Loop
         
         mrsSearchedProviders.MoveFirst
                    
         'Select first provider
         If mcboProviders.ListCount <> 0 Then
            mcboProviders.ListIndex = 0
            mrsSearchedProviders.Find "FLD_ORDER = " & mcboProviders.ItemData(0)
            If Not mrsSearchedProviders.EOF Then
               msSelectedProviderIEN = mrsSearchedProviders("IEN")
            Else
               msSelectedProviderIEN = ""
            End If
         End If
           
      End If
   
   End If
        
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".FillProvidersList", Exe, vbExclamation)

End Sub

Public Sub SearchProviders()

On Error GoTo Error_Handler

Dim bRet As Boolean
   
   'Show provider search dialog
   cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION
   bRet = mobjProviderSearch.SearchClinician(mrsSearchedProviders)
   cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ORDER_SETS_APPLICATION
   
   FillProvidersList
   
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".SearchProviders", Exe, vbExclamation)
   
End Sub

Public Sub MoveControls(ByVal lLeft As Long, ByVal lTop As Long)
   
On Error GoTo Error_Handler

   InitCheck
   
   mlblCaption.Top = lTop + CAPTION_VERTICAL_OFFSET
   mcboProviders.Top = lTop
   mcmdSearch.Top = lTop
    
   mlblCaption.Left = lLeft
   mcboProviders.Left = mlblCaption.Left + mlblCaption.Width + COMBO_HORIZONTAL_OFFSET
   mcmdSearch.Left = mcboProviders.Left + mcboProviders.Width + BUTTON_HORIZONTAL_OFFSET
  
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".MoveControls", Exe, vbExclamation)
   
End Sub

Public Sub HideControls()
   InitCheck
   mlblCaption.Visible = False
   mcboProviders.Visible = False
   mcmdSearch.Visible = False
   mbVisible = False
End Sub

Public Sub ShowControls()
   InitCheck
   mlblCaption.Visible = True
   mcboProviders.Visible = True
   mcmdSearch.Visible = True
   mbVisible = True
End Sub

Public Sub EnableControls()
   InitCheck
   mcboProviders.Enabled = True
   mcmdSearch.Enabled = True
   mbEnabled = True
End Sub

Public Sub DisableControls()
   InitCheck
   mcboProviders.Enabled = False
   mcmdSearch.Enabled = False
   mbEnabled = False
End Sub

Private Sub InitCheck()
   If Not mbInit Then
      'Warn developer they need to initialize decorator class first
      Err.Raise 911, , "Init must first be called!"
   End If
End Sub

Property Get IsSignableProvider()
   InitCheck
   IsSignableProvider = mobjProviderSearch.SignableProvider
End Property

Property Get ProviderSearchObject() As CHCSII_NonProvider.clsNonProvider
   'Provided to enable pass through of all NPOE class functionality
   Set ProviderSearchObject = mobjProviderSearch
End Property

Private Sub mcboProviders_Click()
   ProviderSelected
End Sub

Private Sub mcmdSearch_Click()
   SearchProviders
End Sub

Property Get SelectedProviderIEN() As String
   SelectedProviderIEN = msSelectedProviderIEN
End Property

Private Sub ProviderSelected()

On Error GoTo Error_Handler

'The first one is the choice (Index 0) if the ListIndex is not on 0 then run the procedure
            'mbProviderListChanged = True
    
   If mcboProviders.ListIndex <> 0 Then
   
      If mcboProviders.ListIndex <> -1 Then
         'Let the NPOE dll know what the choice is.
         'The Dll will re-arrange the list with the choice at the top (ListIndex 0)
         mobjProviderSearch.ProviderSelected mrsSearchedProviders, mcboProviders.ItemData(mcboProviders.ListIndex)
         
         'Now re-populate the combobox with the approporiate order
         FillProvidersList
                           
      End If
                
      'Now set the selection to the 0 position
      mcboProviders.ListIndex = 0
'      cboRequestingProvider.ValueCol = 1
'      goOESession.OrderingProviderIen = cboRequestingProvider.Value
'      cboRequestingProvider.ValueCol = 0
   End If
        
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".ProviderSelected", Exe, vbExclamation)
   
End Sub

Property Get Visible() As Boolean
   Visible = mbVisible
End Property

Friend Sub RefreshProviders(ByVal sProviderIEN As Long)
'Called when the provider was changed in another so that the list can be refreshed and kept
'in sync with the other modules

On Error GoTo Error_Handler
   
   msSelectedProviderIEN = sProviderIEN
   mobjProviderSearch.getProviders mrsSearchedProviders
   FillProvidersList
   
Error_Handler:
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, MODULE_NAME & ".RefreshProviders", Exe, vbExclamation)
   
End Sub
