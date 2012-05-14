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
Attribute VB_Name = "clsProph"
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
            'Prophylaxis section ***********************************
            If .PRH_PROPH_Complete Or .PRH_PROPH_Incomplete Then
                sNote = sNote + crtfStartControlName & "Dental Prophylaxis: "
                If .PRH_PROPH_Complete Then
                    sNote = sNote & "Complete Prophylaxis; "
                End If
                   
                If .PRH_PROPH_Incomplete Then
                    sNote = sNote & "Incomplete Prophylaxis; "
                End If
                'Remove the last "; "
                sNote = Left(sNote, Len(sNote) - Len("; "))
                sNote = sNote & crtfEndControlName
            End If
            
            'Fluoride used **************************************************
            If .PRH_PROPH_Fluoride_SNF2 Or .PRH_PROPH_Fluoride_APF Or _
                .PRH_PROPH_Fluoride_NAF Or .PRH_PROPH_Fluoride_Varnish Then
                sNote = sNote + crtfStartControlName & "Fluoride used: "
                If .PRH_PROPH_Fluoride_SNF2 Then
                    sNote = sNote & "8% SnF; "
                End If
        
                If .PRH_PROPH_Fluoride_APF Then
                    sNote = sNote & "1.23% APF; "
                End If
        
                If .PRH_PROPH_Fluoride_NAF Then
                    sNote = sNote & "2% NaF; "
                End If
        
                If .PRH_PROPH_Fluoride_Varnish Then
                    sNote = sNote & "5% NaF varnish; "
                End If
                'Remove the last "; "
                sNote = Left(sNote, Len(sNote) - Len("; "))
                sNote = sNote & crtfEndControlName
            End If
    
            
            If .PRH_SUP_FullMouth Or .PRH_SUP_MaxRight Or .PRH_SUP_MaxAnt Or .PRH_SUP_MaxLeft _
                Or .PRH_SUP_ManRight Or .PRH_SUP_ManAnt Or .PRH_SUP_ManLeft Then
                sNote = sNote + crtfStartControlName & "Supragingival scaling: "
                If .PRH_SUP_FullMouth Then
                    sNote = sNote & "Full mouth supragingival scaling; "
                End If
        
                If .PRH_SUP_MaxRight Then
                    sNote = sNote & "Max right sextant; "
                End If
        
                If .PRH_SUP_MaxAnt Then
                    sNote = sNote & "Max ant sextant; "
                End If
        
                If .PRH_SUP_MaxLeft Then
                    sNote = sNote & "Max left sextant; "
                End If
        
                If .PRH_SUP_ManRight Then
                    sNote = sNote & "Man right sextant; "
                End If
        
                If .PRH_SUP_ManAnt Then
                    sNote = sNote & "Man ant sextant; "
                End If
        
                If .PRH_SUP_ManLeft Then
                    sNote = sNote & "Man left sextant; "
                End If
                'Remove the last "; "
                sNote = Left(sNote, Len(sNote) - Len("; "))
                sNote = sNote & crtfEndControlName
            End If
    
            'Polish ************************************************************
            If .PRH_POL_FullMouth Or .PRH_POL_MaxRight Or .PRH_POL_MaxAnt Or .PRH_POL_MaxLeft _
                Or .PRH_POL_ManRight Or .PRH_POL_ManAnt Or .PRH_POL_ManLeft Then
                sNote = sNote + crtfStartControlName & "Polish: "
                If .PRH_POL_FullMouth Then
                    sNote = sNote & "Full mouth polished all quadrants; "
                End If
        
                If .PRH_POL_MaxRight Then
                    sNote = sNote & "Max right sextant; "
                End If
        
                If .PRH_POL_MaxAnt Then
                    sNote = sNote & "Max ant sextant; "
                End If
        
                If .PRH_POL_MaxLeft Then
                    sNote = sNote & "Max left sextant; "
                End If
        
                If .PRH_POL_ManRight Then
                    sNote = sNote & "Man right sextant; "
                End If
        
                If .PRH_POL_ManAnt Then
                    sNote = sNote & "Man ant sextant; "
                End If
        
                If .PRH_POL_ManLeft Then
                    sNote = sNote & "Man left sextant; "
                End If
                'Remove the last "; "
                sNote = Left(sNote, Len(sNote) - Len("; "))
                sNote = sNote & crtfEndControlName
            End If
            
                  
            'Prophylaxis Notes
            If .PRH_Prophylaxis_Comments <> "" Then
                sNote = sNote + crtfStartControlName & "Prophylaxis Comment: "
                sNote = sNote & .PRH_Prophylaxis_Comments
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
