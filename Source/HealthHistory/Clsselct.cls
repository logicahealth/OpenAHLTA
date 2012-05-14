VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const INVERSE = 6

Public fLeft As Single
Public fTop As Single
Public fRight As Single
Public fBottom As Single
Public fHorizontalMid As Single
Public fVerticalMid As Single
Public ctrControl As Control
Public mbDrawDone As Boolean

Public mbMoveDone As Boolean
Public fXPoint As Single
Public fYPoint As Single
Public fMoveLeft As Single
Public fMoveTop As Single
Public fMoveRight As Single
Public fMoveBottom As Single
Public ctrMoveControl As Control

Public mbResizeDone As Boolean
Public mnResizeMode As Integer
Public mbResizeOutside As Integer
Public mfResizeY As Single
Public mfResizeX As Single

Public frmParent As Form  'added 8/29/02 by Scott Lehman... need a "hook" to help with refreshing

Public Sub DrawSelectedBox(ctrSelectedControl As Control)
    Dim nSavedMode As Integer
    Dim lSuccess As Long
    
    Set ctrControl = ctrSelectedControl
    
    nSavedMode = ctrControl.Container.DrawMode
    ctrControl.Container.DrawMode = vbBlackness
    
    fLeft = ctrSelectedControl.Left
    fTop = ctrSelectedControl.Top
    fRight = fLeft + ctrSelectedControl.Width
    fBottom = fTop + ctrSelectedControl.Height
    fHorizontalMid = fLeft + (ctrSelectedControl.Width / 2)
    fVerticalMid = fTop + (ctrSelectedControl.Height / 2)
            
    'Top Boxes
    ctrControl.Container.Line (fLeft - 30, fTop - 30)-(fLeft + 10, fTop + 10), , BF
    ctrControl.Container.Line (fHorizontalMid - 20, fTop - 30)-(fHorizontalMid + 20, fTop + 10), , BF
    ctrControl.Container.Line (fRight - 10, fTop - 30)-(fRight + 30, fTop + 10), , BF
    
    'Side Boxes
    ctrControl.Container.Line (fLeft - 30, fVerticalMid - 20)-(fLeft + 10, fVerticalMid + 20), , BF
    ctrControl.Container.Line (fRight - 10, fVerticalMid - 20)-(fRight + 30, fVerticalMid + 20), , BF
    
    'Bottom Boxes
    ctrControl.Container.Line (fLeft - 30, fBottom + 30)-(fLeft + 10, fBottom - 10), , BF
    ctrControl.Container.Line (fHorizontalMid - 20, fBottom + 30)-(fHorizontalMid + 20, fBottom - 10), , BF
    ctrControl.Container.Line (fRight - 10, fBottom + 30)-(fRight + 30, fBottom - 10), , BF
    
    ctrControl.Container.DrawMode = nSavedMode
    fXPoint = 0
    fYPoint = 0
    
    mbDrawDone = True
End Sub

Public Sub ErasePrevBox()
    Dim nSavedMode As Integer
    
    If (mbDrawDone = True) Then
        nSavedMode = ctrControl.Container.DrawMode
        ctrControl.Container.DrawMode = vbCopyPen
    
        'Top Boxes
        ctrControl.Container.Line (fLeft - 30, fTop - 30)-(fLeft + 10, fTop + 10), , BF
        ctrControl.Container.Line (fHorizontalMid - 20, fTop - 30)-(fHorizontalMid + 20, fTop + 10), , BF
        ctrControl.Container.Line (fRight - 10, fTop - 30)-(fRight + 30, fTop + 10), , BF
        
        'Side Boxes
        ctrControl.Container.Line (fLeft - 30, fVerticalMid - 20)-(fLeft + 10, fVerticalMid + 20), , BF
        ctrControl.Container.Line (fRight - 10, fVerticalMid - 20)-(fRight + 30, fVerticalMid + 20), , BF
        
        'Bottom Boxes
        ctrControl.Container.Line (fLeft - 30, fBottom - 10)-(fLeft + 10, fBottom + 30), , BF
        ctrControl.Container.Line (fHorizontalMid - 20, fBottom - 10)-(fHorizontalMid + 20, fBottom + 30), , BF
        ctrControl.Container.Line (fRight - 10, fBottom - 10)-(fRight + 30, fBottom + 30), , BF
        
        ctrControl.Container.DrawMode = nSavedMode
        ctrControl.Container.Refresh
        mbDrawDone = False
    End If
    
    '08-28-02 - SCR #26989 - Scott Lehman
    '   need to erase prior tracks of the moved box
    Dim oConf As clsConfigModule
    For Each oConf In frmProperties.mcolConfMod
        oConf.LblBox.Refresh
    Next
End Sub
Public Sub MoveSelectedBox(X As Single, Y As Single, ctrSelectedControl As Control)
    Dim nSavedMode As Integer
    Dim nSavedWidth As Integer
    Dim fHorizontalChange As Single
    Dim fVerticalChange As Single
    
    If (Not (ctrMoveControl Is Nothing)) Then
        If (ctrMoveControl <> ctrSelectedControl) Then
            fXPoint = 0
            fYPoint = 0
        End If
    End If
    Set ctrMoveControl = ctrSelectedControl
    nSavedMode = ctrMoveControl.Container.DrawMode
    nSavedWidth = ctrMoveControl.Container.DrawWidth
    ctrMoveControl.Container.DrawWidth = 2
    
    
    If (fXPoint <> 0) Or (fYPoint <> 0) Then
        'Calculate the change
        fHorizontalChange = X - fXPoint
        fVerticalChange = Y - fYPoint
        If (fHorizontalChange > 20) Or (fHorizontalChange < -20) Or (fVerticalChange > 20) Or (fVerticalChange < -20) Then
            ctrMoveControl.Container.DrawMode = vbCopyPen
            'Erase the Previous Box
            ctrMoveControl.Container.Line (fMoveLeft, fMoveTop)-(fMoveRight, fMoveBottom), , B
            'Check to see if we need to refresh
            If (fMoveLeft >= ctrMoveControl.Left) And (fMoveLeft <= ctrMoveControl.Left + ctrMoveControl.Width) And (fMoveTop >= ctrMoveControl.Top) And (fMoveTop <= ctrMoveControl.Top + ctrMoveControl.Height) Then
                ctrMoveControl.Container.Refresh
            ElseIf (fMoveLeft >= ctrMoveControl.Left) And (fMoveLeft <= ctrMoveControl.Left + ctrMoveControl.Width) And (fMoveBottom >= ctrMoveControl.Top) And (fMoveBottom <= ctrMoveControl.Top + ctrMoveControl.Height) Then
                ctrMoveControl.Container.Refresh
            ElseIf (fMoveRight >= ctrMoveControl.Left) And (fMoveRight <= ctrMoveControl.Left + ctrMoveControl.Width) And (fMoveTop >= ctrMoveControl.Top) And (fMoveTop <= ctrMoveControl.Top + ctrMoveControl.Height) Then
                ctrMoveControl.Container.Refresh
            ElseIf (fMoveRight >= ctrMoveControl.Left) And (fMoveRight <= ctrMoveControl.Left + ctrMoveControl.Width) And (fMoveBottom >= ctrMoveControl.Top) And (fMoveBottom <= ctrMoveControl.Top + ctrMoveControl.Height) Then
                ctrMoveControl.Container.Refresh
            End If
            ctrMoveControl.Container.DrawMode = vbWhiteness
            'Draw the New Box
            fMoveLeft = fMoveLeft + fHorizontalChange
            fMoveTop = fMoveTop + fVerticalChange
            fMoveRight = fMoveRight + fHorizontalChange
            fMoveBottom = fMoveBottom + fVerticalChange
            
            'Make sure we are within the boundries
            If (fMoveLeft < 0) Then
                fMoveRight = fMoveRight - fMoveLeft
                fMoveLeft = 0
            ElseIf (fMoveRight > ctrMoveControl.Container.Width) Then
                fMoveLeft = fMoveLeft - (fMoveRight - ctrMoveControl.Container.Width)
                fMoveRight = ctrMoveControl.Container.Width
            End If
            If (fMoveTop < 0) Then
                fMoveBottom = fMoveBottom - fMoveTop
                fMoveTop = 0
            ElseIf (fMoveBottom > ctrMoveControl.Container.Height) Then
                fMoveTop = fMoveTop - (fMoveBottom - ctrMoveControl.Container.Height)
                fMoveBottom = ctrMoveControl.Container.Height
            End If
        
            ctrMoveControl.Container.Line (fMoveLeft, fMoveTop)-(fMoveRight, fMoveBottom), , B
            fXPoint = X
            fYPoint = Y
        End If
    Else
        'First Time so just draw the Box using the controls coordinates
        fMoveLeft = ctrMoveControl.Left
        fMoveTop = ctrMoveControl.Top
        fMoveRight = fMoveLeft + ctrMoveControl.Width
        fMoveBottom = fMoveTop + ctrMoveControl.Height
        ctrMoveControl.Container.DrawMode = vbWhiteness
        ctrMoveControl.Container.Line (fMoveLeft, fMoveTop)-(fMoveRight, fMoveBottom), , B
        fXPoint = X
        fYPoint = Y
    End If
    
    
    ctrMoveControl.Container.DrawMode = nSavedMode
    ctrMoveControl.Container.DrawWidth = nSavedWidth
    mbMoveDone = True
    
    '08-28-02 - SCR #26989 - Scott Lehman
    '   need to erase prior tracks of the moved box
    ErasePrevBox
End Sub
Public Sub MoveControlToMoveCoordinates(ctrSelectedControl As Control)
    Dim nSavedMode As Integer
    If (mbMoveDone) Then
        nSavedMode = ctrSelectedControl.Container.DrawMode
        ctrSelectedControl.Container.DrawMode = vbCopyPen
        'Erase the Previous Box
        ctrSelectedControl.Container.Line (fMoveLeft, fMoveTop)-(fMoveRight, fMoveBottom), , B
        ctrSelectedControl.Left = fMoveLeft
        ctrSelectedControl.Top = fMoveTop
        ctrSelectedControl.Container.Refresh
        fXPoint = 0
        fYPoint = 0
        ctrSelectedControl.Container.DrawMode = nSavedMode
        mbMoveDone = False
    End If
End Sub
Public Function SetMousePointer(MyForm As Form, vfX As Single, vfY As Single) As Boolean
    If (vfX >= (fLeft - 40)) And (vfX <= fLeft + 20) And (vfY >= fTop - 40) And (vfY <= fTop + 20) Then
        MyForm.MousePointer = vbSizeNWSE
        SetMousePointer = True
        mnResizeMode = 1
    ElseIf (vfX >= (fRight - 20)) And (vfX <= fRight + 40) And (vfY >= fBottom - 20) And (vfY <= fBottom + 40) Then
        MyForm.MousePointer = vbSizeNWSE
        SetMousePointer = True
        mnResizeMode = 2
    ElseIf (vfX >= (fRight - 20)) And (vfX <= fRight + 40) And (vfY >= fTop - 40) And (vfY <= fTop + 20) Then
        MyForm.MousePointer = vbSizeNESW
        SetMousePointer = True
        mnResizeMode = 3
    ElseIf (vfX >= (fLeft - 40)) And (vfX <= fLeft + 20) And (vfY >= fBottom - 20) And (vfY <= fBottom + 40) Then
        MyForm.MousePointer = vbSizeNESW
        SetMousePointer = True
        mnResizeMode = 4
    ElseIf (vfX >= (fHorizontalMid - 30)) And (vfX <= fHorizontalMid + 30) And (vfY >= fTop - 40) And (vfY <= fTop + 20) Then
        MyForm.MousePointer = vbSizeNS
        SetMousePointer = True
        mnResizeMode = 5
    ElseIf (vfX >= (fHorizontalMid - 30)) And (vfX <= fHorizontalMid + 30) And (vfY >= fBottom - 20) And (vfY <= fBottom + 40) Then
        MyForm.MousePointer = vbSizeNS
        SetMousePointer = True
        mnResizeMode = 6
    ElseIf (vfX >= (fLeft - 40)) And (vfX <= fLeft + 20) And (vfY >= fVerticalMid - 30) And (vfY <= fVerticalMid + 30) Then
        MyForm.MousePointer = vbSizeWE
        SetMousePointer = True
        mnResizeMode = 7
    ElseIf (vfX >= (fRight - 20)) And (vfX <= fRight + 40) And (vfY >= fVerticalMid - 30) And (vfY <= fVerticalMid + 30) Then
        MyForm.MousePointer = vbSizeWE
        SetMousePointer = True
        mnResizeMode = 8
    Else
        MyForm.MousePointer = vbDefault
        SetMousePointer = False
        mnResizeMode = 0
    End If
End Function
Public Sub ResizeSelectedBox(vfX As Single, vfY As Single)
    Dim nSavedMode As Integer
    Dim nSavedWidth As Integer
       
    nSavedMode = ctrControl.Container.DrawMode
    nSavedWidth = ctrControl.Container.DrawWidth
    
    ctrControl.Container.DrawWidth = 2
        
    Select Case mnResizeMode
        Case 1               'This is the Top-Left corner
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 4 Lines
                ctrControl.Container.Line (fRight, fTop)-(fRight, mfResizeY)
                ctrControl.Container.Line (mfResizeX, mfResizeY)-(fRight, mfResizeY)
                ctrControl.Container.Line (mfResizeX, mfResizeY)-(mfResizeX, fBottom)
                ctrControl.Container.Line (mfResizeX, fBottom)-(fLeft, fBottom)
            End If
            ctrControl.Container.Refresh
            ctrControl.Container.DrawMode = vbWhiteness
            If (vfY > fBottom) Then
                mfResizeY = fBottom
            ElseIf (vfY > 0) Then
                mfResizeY = vfY
            Else
                mfResizeY = 0
            End If
            If (vfX > fRight) Then
                mfResizeX = fRight
            ElseIf (vfX > 0) Then
                mfResizeX = vfX
            Else
                mfResizeX = 0
            End If
            'Need to draw 4 lines
            ctrControl.Container.Line (fRight, fTop)-(fRight, mfResizeY)
            ctrControl.Container.Line (mfResizeX, mfResizeY)-(fRight, mfResizeY)
            ctrControl.Container.Line (mfResizeX, mfResizeY)-(mfResizeX, fBottom)
            ctrControl.Container.Line (mfResizeX, fBottom)-(fLeft, fBottom)
            mbResizeOutside = True
            mbResizeDone = True
            
        Case 2                      'This is the Bottom-Right corner
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 4 Lines
                ctrControl.Container.Line (fRight, fTop)-(mfResizeX, fTop)
                ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, mfResizeY)
                ctrControl.Container.Line (mfResizeX, mfResizeY)-(fLeft, mfResizeY)
                ctrControl.Container.Line (fLeft, fBottom)-(fLeft, mfResizeY)
            End If
            ctrControl.Container.Refresh
            ctrControl.Container.DrawMode = vbWhiteness
            If (vfY < fTop) Then
                mfResizeY = fTop
            ElseIf (vfY < ctrControl.Container.Height) Then
                mfResizeY = vfY
            Else
                mfResizeY = ctrControl.Container.Height
            End If
            If (vfX < fLeft) Then
                mfResizeX = fLeft
            ElseIf (vfX < ctrControl.Container.Width) Then
                mfResizeX = vfX
            Else
                mfResizeX = ctrControl.Container.Width
            End If
            'Need to draw 4 lines
            ctrControl.Container.Line (fRight, fTop)-(mfResizeX, fTop)
            ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, mfResizeY)
            ctrControl.Container.Line (mfResizeX, mfResizeY)-(fLeft, mfResizeY)
            ctrControl.Container.Line (fLeft, fBottom)-(fLeft, mfResizeY)
            mbResizeOutside = True
            mbResizeDone = True
            
        Case 3                      'This is the Top-Right Corner
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 4 Lines
                ctrControl.Container.Line (fLeft, fTop)-(fLeft, mfResizeY)
                ctrControl.Container.Line (fLeft, mfResizeY)-(mfResizeX, mfResizeY)
                ctrControl.Container.Line (mfResizeX, mfResizeY)-(mfResizeX, fBottom)
                ctrControl.Container.Line (mfResizeX, fBottom)-(fRight, fBottom)
            End If
            ctrControl.Container.Refresh
            ctrControl.Container.DrawMode = vbWhiteness
            If (vfY > fBottom) Then
                mfResizeY = fBottom
            ElseIf (vfY > 0) Then
                mfResizeY = vfY
            Else
                mfResizeY = 0
            End If
            If (vfX < fLeft) Then
                mfResizeX = fLeft
            ElseIf (vfX < ctrControl.Container.Width) Then
                mfResizeX = vfX
            Else
                mfResizeX = ctrControl.Container.Width
            End If
            'Need to draw 4 lines
            ctrControl.Container.Line (fLeft, fTop)-(fLeft, mfResizeY)
            ctrControl.Container.Line (fLeft, mfResizeY)-(mfResizeX, mfResizeY)
            ctrControl.Container.Line (mfResizeX, mfResizeY)-(mfResizeX, fBottom)
            ctrControl.Container.Line (mfResizeX, fBottom)-(fRight, fBottom)
            mbResizeOutside = True
            mbResizeDone = True
            
        Case 4                       'This is the Bottom-Left Corner
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 4 Lines
                ctrControl.Container.Line (fLeft, fTop)-(mfResizeX, fTop)
                ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, mfResizeY)
                ctrControl.Container.Line (mfResizeX, mfResizeY)-(fRight, mfResizeY)
                ctrControl.Container.Line (fRight, mfResizeY)-(fRight, fBottom)
            End If
            ctrControl.Container.Refresh
            ctrControl.Container.DrawMode = vbWhiteness
            If (vfY < fTop) Then
                mfResizeY = fTop
            ElseIf (vfY < ctrControl.Container.Height) Then
                mfResizeY = vfY
            Else
                mfResizeY = ctrControl.Container.Height
            End If
            If (vfX > fRight) Then
                mfResizeX = fRight
            ElseIf (vfX > 0) Then
                mfResizeX = vfX
            Else
                mfResizeX = 0
            End If
            'Need to draw 4 lines
            ctrControl.Container.Line (fLeft, fTop)-(mfResizeX, fTop)
            ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, mfResizeY)
            ctrControl.Container.Line (mfResizeX, mfResizeY)-(fRight, mfResizeY)
            ctrControl.Container.Line (fRight, mfResizeY)-(fRight, fBottom)
            mbResizeOutside = True
            mbResizeDone = True
            
        Case 5                      'This is the Top Middle
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 3 Lines
                ctrControl.Container.Line (fLeft, fTop)-(fLeft, mfResizeY)
                ctrControl.Container.Line (fRight, fTop)-(fRight, mfResizeY)
                ctrControl.Container.Line (fLeft, mfResizeY)-(fRight, mfResizeY)
            Else
                ctrControl.Container.Refresh
            End If
            ctrControl.Container.DrawMode = vbWhiteness
            If (vfY < fTop) Then
                If (vfY < 0) Then
                    vfY = 0
                End If
                'Need to draw 3 lines
                ctrControl.Container.Line (fLeft, fTop)-(fLeft, vfY)
                ctrControl.Container.Line (fRight, fTop)-(fRight, vfY)
                ctrControl.Container.Line (fLeft, vfY)-(fRight, vfY)
                mfResizeY = vfY
                mbResizeOutside = True
            ElseIf (vfY >= fTop) And (vfY <= fBottom) Then
                ctrControl.Container.Line (fLeft, vfY)-(fRight, vfY)
                mfResizeY = vfY
                mbResizeOutside = False
            Else
                ctrControl.Container.Line (fLeft, fBottom)-(fRight, fBottom)
                mfResizeY = fBottom
                mbResizeOutside = False
            End If
            mbResizeDone = True
            
        Case 6                   'This is the Bottom middle
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 3 Lines
                ctrControl.Container.Line (fLeft, fBottom)-(fLeft, mfResizeY)
                ctrControl.Container.Line (fRight, fBottom)-(fRight, mfResizeY)
                ctrControl.Container.Line (fLeft, mfResizeY)-(fRight, mfResizeY)
            Else
                ctrControl.Container.Refresh
            End If
            ctrControl.Container.DrawMode = vbWhiteness
            
            If (vfY > fBottom) Then
                If (vfY > ctrControl.Container.Height) Then
                    vfY = ctrControl.Container.Height
                End If
                'Need to draw 3 lines
                ctrControl.Container.Line (fLeft, fBottom)-(fLeft, vfY)
                ctrControl.Container.Line (fRight, fBottom)-(fRight, vfY)
                ctrControl.Container.Line (fLeft, vfY)-(fRight, vfY)
                mfResizeY = vfY
                mbResizeOutside = True
            ElseIf (vfY <= fBottom) And (vfY >= fTop) Then
                ctrControl.Container.Line (fLeft, vfY)-(fRight, vfY)
                mfResizeY = vfY
                mbResizeOutside = False
            Else
                ctrControl.Container.Line (fLeft, fTop)-(fRight, fTop)
                mfResizeY = fTop
                mbResizeOutside = False
            End If
            mbResizeDone = True
            
        Case 7                  'This is Middle Left
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 3 Lines
                ctrControl.Container.Line (mfResizeX, fTop)-(fLeft, fTop)
                ctrControl.Container.Line (mfResizeX, fBottom)-(fLeft, fBottom)
                ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, fBottom)
            Else
                ctrControl.Container.Refresh
            End If
            ctrControl.Container.DrawMode = vbWhiteness
            
            If (vfX < fLeft) Then
                If (vfX < 0) Then
                    vfX = 0
                End If
                'Need to draw 3 lines
                ctrControl.Container.Line (vfX, fTop)-(fLeft, fTop)
                ctrControl.Container.Line (vfX, fBottom)-(fLeft, fBottom)
                ctrControl.Container.Line (vfX, fTop)-(vfX, fBottom)
                mfResizeX = vfX
                mbResizeOutside = True
            ElseIf (vfX >= fLeft) And (vfX <= fRight) Then
                ctrControl.Container.Line (vfX, fTop)-(vfX, fBottom)
                mfResizeX = vfX
                mbResizeOutside = False
            Else
                ctrControl.Container.Line (fRight, fTop)-(fRight, fBottom)
                mfResizeX = fRight
                mbResizeOutside = False
            End If
            mbResizeDone = True
            
        Case 8               'This is the Right Middle
            If (mbResizeOutside) Then
                ctrControl.Container.DrawMode = vbCopyPen
                'Need to Erase 3 Lines
                ctrControl.Container.Line (mfResizeX, fTop)-(fRight, fTop)
                ctrControl.Container.Line (mfResizeX, fBottom)-(fRight, fBottom)
                ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, fBottom)
            Else
                ctrControl.Container.Refresh
            End If
            ctrControl.Container.DrawMode = vbWhiteness
            If (vfX > fRight) Then
                If (vfX > ctrControl.Container.Width) Then
                    vfX = ctrControl.Container.Width
                End If
                'Need to draw 3 lines
                ctrControl.Container.Line (vfX, fTop)-(fRight, fTop)
                ctrControl.Container.Line (vfX, fBottom)-(fRight, fBottom)
                ctrControl.Container.Line (vfX, fTop)-(vfX, fBottom)
                mfResizeX = vfX
                mbResizeOutside = True
            ElseIf (vfX >= fLeft) And (vfX <= fRight) Then
                ctrControl.Container.Line (vfX, fTop)-(vfX, fBottom)
                mfResizeX = vfX
                mbResizeOutside = False
            Else
                ctrControl.Container.Line (fLeft, fTop)-(fLeft, fBottom)
                mfResizeX = fLeft
                mbResizeOutside = False
            End If
            mbResizeDone = True
        Case Else
            mbResizeDone = False
    End Select
    ctrControl.Container.DrawMode = nSavedMode
    ctrControl.Container.DrawWidth = nSavedWidth
End Sub




Sub MoveControlToResizeCoordinates()
    Dim nSavedMode As Integer
    Dim nSavedWidth As Integer
    Dim sHeightChange As Single
    Dim sWidthChange
       
    If (mbResizeDone) Then
        nSavedMode = ctrControl.Container.DrawMode
        nSavedWidth = ctrControl.Container.DrawWidth
        
        ctrControl.Container.DrawWidth = 2
        
        Select Case mnResizeMode
            Case 1
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 4 Lines
                    ctrControl.Container.Line (fRight, fTop)-(fRight, mfResizeY)
                    ctrControl.Container.Line (mfResizeX, mfResizeY)-(fRight, mfResizeY)
                    ctrControl.Container.Line (mfResizeX, mfResizeY)-(mfResizeX, fBottom)
                    ctrControl.Container.Line (mfResizeX, fBottom)-(fLeft, fBottom)
                End If
                'Resize the control
                If (mfResizeY >= fTop) Then
                    sHeightChange = -(mfResizeY - fTop)
                Else
                    sHeightChange = fTop - mfResizeY
                End If
                If (mfResizeX >= fLeft) Then
                    sWidthChange = -(mfResizeX - fLeft)
                Else
                    sWidthChange = fLeft - mfResizeX
                End If
                ctrControl.Height = ctrControl.Height + sHeightChange
                ctrControl.Width = ctrControl.Width + sWidthChange
                ctrControl.Top = mfResizeY
                ctrControl.Left = mfResizeX
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 2
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 4 Lines
                    ctrControl.Container.Line (fRight, fTop)-(mfResizeX, fTop)
                    ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, mfResizeY)
                    ctrControl.Container.Line (mfResizeX, mfResizeY)-(fLeft, mfResizeY)
                    ctrControl.Container.Line (fLeft, fBottom)-(fLeft, mfResizeY)
                End If
                'Resize the control
                If (mfResizeY <= fBottom) Then
                    sHeightChange = -(fBottom - mfResizeY)
                Else
                    sHeightChange = mfResizeY - fBottom
                End If
                If (mfResizeX <= fRight) Then
                    sWidthChange = -(fRight - mfResizeX)
                Else
                    sWidthChange = mfResizeX - fRight
                End If
                ctrControl.Height = ctrControl.Height + sHeightChange
                ctrControl.Width = ctrControl.Width + sWidthChange
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 3
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 4 Lines
                    ctrControl.Container.Line (fLeft, fTop)-(fLeft, mfResizeY)
                    ctrControl.Container.Line (fLeft, mfResizeY)-(mfResizeX, mfResizeY)
                    ctrControl.Container.Line (mfResizeX, mfResizeY)-(mfResizeX, fBottom)
                    ctrControl.Container.Line (mfResizeX, fBottom)-(fRight, fBottom)
                End If
                'Resize the control
                If (mfResizeY > fTop) Then
                    sHeightChange = -(mfResizeY - fTop)
                Else
                    sHeightChange = fTop - mfResizeY
                End If
                If (mfResizeX < fRight) Then
                    sWidthChange = -(fRight - mfResizeX)
                Else
                    sWidthChange = mfResizeX - fRight
                End If
                ctrControl.Height = ctrControl.Height + sHeightChange
                ctrControl.Width = ctrControl.Width + sWidthChange
                ctrControl.Top = mfResizeY
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 4
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 4 Lines
                    ctrControl.Container.Line (fLeft, fTop)-(mfResizeX, fTop)
                    ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, mfResizeY)
                    ctrControl.Container.Line (mfResizeX, mfResizeY)-(fRight, mfResizeY)
                    ctrControl.Container.Line (fRight, mfResizeY)-(fRight, fBottom)
                End If
                'Resize the control
                If (mfResizeY <= fBottom) Then
                    sHeightChange = -(fBottom - mfResizeY)
                Else
                    sHeightChange = mfResizeY - fBottom
                End If
                If (mfResizeX > fLeft) Then
                    sWidthChange = -(mfResizeX - fLeft)
                Else
                    sWidthChange = fLeft - mfResizeX
                End If
                ctrControl.Height = ctrControl.Height + sHeightChange
                ctrControl.Width = ctrControl.Width + sWidthChange
                ctrControl.Left = mfResizeX
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 5
                'Erase the Lines
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 3 Lines
                    ctrControl.Container.Line (fLeft, fTop)-(fLeft, mfResizeY)
                    ctrControl.Container.Line (fRight, fTop)-(fRight, mfResizeY)
                    ctrControl.Container.Line (fLeft, mfResizeY)-(fRight, mfResizeY)
                End If
                'Resize the control
                If (mfResizeY >= fTop) Then
                    sHeightChange = -(mfResizeY - fTop)
                Else
                    sHeightChange = fTop - mfResizeY
                End If
                ctrControl.Height = ctrControl.Height + sHeightChange
                fTop = mfResizeY
                ctrControl.Top = mfResizeY
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 6
                'Erase the Lines
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 3 Lines
                    ctrControl.Container.Line (fLeft, fBottom)-(fLeft, mfResizeY)
                    ctrControl.Container.Line (fRight, fBottom)-(fRight, mfResizeY)
                    ctrControl.Container.Line (fLeft, mfResizeY)-(fRight, mfResizeY)
                End If
                'Resize the control
                If (mfResizeY <= fBottom) Then
                    sHeightChange = -(fBottom - mfResizeY)
                Else
                    sHeightChange = mfResizeY - fBottom
                End If
                ctrControl.Height = ctrControl.Height + sHeightChange
                fBottom = mfResizeY
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 7
                'Erase the Lines
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 3 Lines
                    ctrControl.Container.Line (mfResizeX, fTop)-(fLeft, fTop)
                    ctrControl.Container.Line (mfResizeX, fBottom)-(fLeft, fBottom)
                    ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, fBottom)
                End If
                'Resize the control
                If (mfResizeX >= fLeft) Then
                    sWidthChange = -(mfResizeX - fLeft)
                Else
                    sWidthChange = fLeft - mfResizeX
                End If
                ctrControl.Width = ctrControl.Width + sWidthChange
                fLeft = mfResizeX
                ctrControl.Left = mfResizeX
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case 8
                'Erase the Lines
                If (mbResizeOutside) Then
                    ctrControl.Container.DrawMode = vbCopyPen
                    'Need to Erase 3 Lines
                    ctrControl.Container.Line (mfResizeX, fTop)-(fRight, fTop)
                    ctrControl.Container.Line (mfResizeX, fBottom)-(fRight, fBottom)
                    ctrControl.Container.Line (mfResizeX, fTop)-(mfResizeX, fBottom)
                End If
                'Resize the control
                If (mfResizeX <= fRight) Then
                    sWidthChange = -(fRight - mfResizeX)
                Else
                    sWidthChange = mfResizeX - fRight
                End If
                ctrControl.Width = ctrControl.Width + sWidthChange
                fRight = mfResizeX
                ctrControl.Refresh
                Call DrawSelectedBox(ctrControl)
            Case Else
        End Select
        mnResizeMode = 0
        mbResizeOutside = False
        mbResizeDone = False
        ctrControl.Container.DrawMode = nSavedMode
        ctrControl.Container.DrawWidth = nSavedWidth
    End If
End Sub
Private Sub Class_Initialize()
    mbDrawDone = False
    mbMoveDone = False
    mnResizeMode = 0
    mbResizeOutside = False
    mbResizeDone = False
    fXPoint = 0
    fYPoint = 0
End Sub

Private Sub Class_Terminate()
    Set ctrControl = Nothing
End Sub
