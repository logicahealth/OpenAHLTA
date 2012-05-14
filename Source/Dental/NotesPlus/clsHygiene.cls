VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHygiene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements clsISection

Private oPreventive As clsPreventive

Public Property Get Preventive() As clsPreventive
    Set Preventive = oPreventive
End Property
Public Property Set Preventive(ByVal objObject As clsPreventive)
    Set oPreventive = objObject
End Property


Private Property Get clsISection_DentalEncounterID() As Long
'
End Property

Private Function clsISection_GetRTFNote() As String
Dim sNote As String
Const sPROC_NAME = "clsISection_GetRTFNote"
On Error GoTo ErrHnd
    If Not oPreventive Is Nothing Then
        
    With oPreventive
         'scaling and Root Planing section ***********************************************************
         If .HG_Scaling_MaxRight_1To3 Or .HG_Scaling_MaxLeft_1To3 Or _
             .HG_Scaling_ManRight_1To3 Or .HG_Scaling_ManLeft_1To3 Or _
             .HG_Scaling_MaxRight_4Plus Or .HG_Scaling_MaxLeft_4Plus Or _
             .HG_Scaling_ManRight_4Plus Or .HG_Scaling_ManLeft_4Plus Then
             
             sNote = sNote + crtfStartControlName & "Scaling and Root Planing: "
             
             If .HG_Scaling_MaxRight_1To3 Then
                 sNote = sNote & "Max right quadrant 1-3 teeth; "
             End If
            
            '1-3 teeth
             If .HG_Scaling_MaxLeft_1To3 Then
                 sNote = sNote & "Max left quadrant 1-3 teeth; "
             End If
                
             If .HG_Scaling_ManRight_1To3 Then
                 sNote = sNote & "Man right quadrant 1-3 teeth; "
             End If
                
             If .HG_Scaling_ManLeft_1To3 Then
                 sNote = sNote & "Man left quadrant 1-3 teeth; "
             End If
             
             '4+ teeth
             If .HG_Scaling_MaxRight_4Plus Then
                 sNote = sNote & "Max right quadrant 4+ teeth; "
             End If
                
             If .HG_Scaling_MaxLeft_4Plus Then
                 sNote = sNote & "Max left quadrant 4+ teeth; "
             End If
                
             If .HG_Scaling_ManRight_4Plus Then
                 sNote = sNote & "Man right quadrant 4+ teeth; "
             End If
                
             If .HG_Scaling_ManLeft_4Plus Then
                 sNote = sNote & "Man left quadrant 4+ teeth; "
             End If
             
             'Remove the last "; "
             sNote = Left(sNote, Len(sNote) - Len("; "))
             sNote = sNote & crtfEndControlName
         End If
         
         'Debridement and Maintenance *************************************************
         If .HG_DM_FullMouth Or .HG_DM_Perio Or .HG_DM_Implant Or .HG_DM_Note <> "" Then
            sNote = sNote + crtfStartControlName & "Debridement and maintenance: "
            If .HG_DM_FullMouth Then
                sNote = sNote & "Full mouth debridement; "
            End If
               
            If .HG_DM_Perio Then
                sNote = sNote & "Periodontal maintenance; "
            End If
               
            If .HG_DM_Implant Then
                sNote = sNote & "Implant maintenance; "
            End If
            
            If .HG_DM_Note <> "" Then
                sNote = sNote & "Debridement and Maintenance Comment: " & .HG_DM_Note & "; "
            End If
            
             'Remove the last "; "
             sNote = Left(sNote, Len(sNote) - Len("; "))
             sNote = sNote & crtfEndControlName
          End If
          
         'Medicaments used - Topical ****************************************************
         If .HG_MED_Chlorhexidine Or .HG_MED_Doxycycline Or .HG_MED_Fluoride_Sodium Or _
              .HG_MED_Fluoride_Stannous Or .HG_MED_Hydrogen Or .HG_MED_Painfree Or _
               .HG_MED_Saline Or .HG_MED_Tetracycline Or .HG_MED_Topical Then
            sNote = sNote + crtfStartControlName & "Medicaments used/Topical: "
            If .HG_MED_Chlorhexidine Then
                sNote = sNote & "Chlorhexidine gluconate; "
            End If
               
             If .HG_MED_Doxycycline Then
                sNote = sNote & "Doxycycline solution; "
            End If
              
            If .HG_MED_Fluoride_Sodium Then
                sNote = sNote & "Fluoride solution (sodium); "
            End If
              
            If .HG_MED_Fluoride_Stannous Then
                sNote = sNote & "Fluoride solution (stannous); "
            End If
              
            If .HG_MED_Hydrogen Then
                sNote = sNote & "Hydrogen peroxide; "
            End If
              
            If .HG_MED_Painfree Then
                sNote = sNote & "pain-free; "
            End If
              
             If .HG_MED_Saline Then
                sNote = sNote & "saline; "
            End If
             
            If .HG_MED_Tetracycline Then
                sNote = sNote & "tetracycline solution; "
            End If
              
            If .HG_MED_Topical Then
                sNote = sNote & "topical anesthetic; "
            End If
             'Remove the last "; "
             sNote = Left(sNote, Len(sNote) - Len("; "))
             sNote = sNote & crtfEndControlName
          End If
           
         'Medicaments used - Site specific ***************************************
        If .HG_MED_Actisite Or .HG_MED_Atridox Or .HG_MED_Perio Then
             sNote = sNote + crtfStartControlName & "Medicaments used/Site specific: "
    
             If .HG_MED_Actisite Then
                 sNote = sNote & "Actisite; "
             End If
                
             If .HG_MED_Atridox Then
                 sNote = sNote & "Atridox; "
             End If
                
             If .HG_MED_Perio Then
                 sNote = sNote & "PerioChip; "
             End If
             'Remove the last "; "
             sNote = Left(sNote, Len(sNote) - Len("; "))
             sNote = sNote & crtfEndControlName
        End If
            
         'Medicaments used - Desensitizing agents ***************************************
         If .HG_MED_Potassium Or .HG_MED_Fluoride_Varnish Or .HG_MED_Duraphet Or .HG_MED_Potassium_Ox Then
             sNote = sNote + crtfStartControlName & "Medicaments used/Desensitizing agents: "
            If .HG_MED_Potassium Then
                sNote = sNote & "Potassium nitrate; "
            End If
            
            'KDunne 4/25/2002
            'SCR 23373
            If .HG_MED_Potassium_Ox Then
                sNote = sNote & "Potassium oxalate; "
            End If
            
            If .HG_MED_Fluoride_Varnish Then
                sNote = sNote & "Fluoride varnish; "
            End If
                                  
            If .HG_MED_Duraphet Then
                sNote = sNote & "Duraphet; "
            End If
             'Remove the last "; "
             sNote = Left(sNote, Len(sNote) - Len("; "))
             sNote = sNote & crtfEndControlName
          End If
          
          'Hygiene Notes ***************************************************************
          If .HG_DM_Note <> "" Or .HG_MED_Note <> "" Then
                 sNote = sNote + crtfStartControlName & "Medicaments Used Note : "
                 sNote = sNote & "Medicaments Used Comment: " & .HG_MED_Note
                 sNote = sNote & crtfEndControlName
          End If
        
        End With
        
    End If

    clsISection_GetRTFNote = sNote
Exit Function

ErrHnd:
    goCWShared.ShowVBError Err.Number, Err.Description, Err.Source & "." & sPROC_NAME, App.Title, vbCritical

End Function

Private Function clsISection_GetNote() As String
    clsISection_GetNote = "*****GetNote Method no longer valid.*****"
End Function


Private Function clsISection_LoadSection(ByVal PatientID As Long, ByVal DentalEncounterID As Long) As Boolean
'
End Function

Private Property Get clsISection_PatientID() As Long
'
End Property

Private Function clsISection_SaveSection() As Boolean
'
End Function

