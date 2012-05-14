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
Attribute VB_Name = "SelectTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a Selection Tab"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Columns"
Attribute VB_Ext_KEY = "Member1" ,"Items"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"ListItems"
Option Explicit

Private mvarColumns As Columns
Private mvarItems As Items
'local variable(s) to hold property value(s)
Private mvarCaption As String 'local copy
Private mvarListItems As ListItems
Private mvarTabIndex As Long

Private mTabIndex As Long
Private mUserControl As TabOne
'local variable(s) to hold property value(s)
Private mvarVisible As Boolean 'local copy


Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mUserControl.TabCaption(mTabIndex) = vData
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mUserControl.TabCaption(mTabIndex)
End Property

Public Property Let Visible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Visible = 5
    Dim i As Long
    
    With mUserControl
    
        .TabVisible(mTabIndex) = vData
        
        If vData = False Then
            For i = 0 To .NumTabs - 1
                If .TabVisible(i) Then
                    .CurrTab = i
                    Exit For
                End If
            Next
        End If
        
    End With
    
End Property

Public Property Get Visible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mUserControl.TabVisible(mTabIndex)
    
End Property



Friend Sub Init(NewUserControl As TabOne, TabIndex As Long)
    Set mUserControl = NewUserControl
    mTabIndex = TabIndex
End Sub


Public Property Get ListItems() As ListItems
    If mvarListItems Is Nothing Then
        Set mvarListItems = New ListItems
        mvarListItems.Init Me
    End If


    Set ListItems = mvarListItems
End Property


Public Property Set ListItems(vData As ListItems)
    Set mvarListItems = vData
End Property





Public Property Get Items() As Items
    If mvarItems Is Nothing Then
        Set mvarItems = New Items
    End If


    Set Items = mvarItems
End Property


Public Property Set Items(vData As Items)
    Set mvarItems = vData
End Property



Public Property Get Columns() As Columns
    If mvarColumns Is Nothing Then
        Set mvarColumns = New Columns
    End If


    Set Columns = mvarColumns
End Property


Public Property Set Columns(vData As Columns)
    Set mvarColumns = vData
End Property
Private Sub Class_Terminate()
  Set mvarListItems = Nothing
  Set mvarItems = Nothing
    Set mvarColumns = Nothing
End Sub
Public Sub Refresh()
    MsgBxARMd "Not Supported At This Time"
End Sub

Public Property Set Picture(vData As IPictureDisp)
    Set mUserControl.TabPicture(mTabIndex) = vData
End Property
