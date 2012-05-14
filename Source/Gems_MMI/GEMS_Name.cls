VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GEMS_Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarLast_name          As String
Private mvarFirst_name         As String
Private mvarMiddle_name        As String
Private mvarSuffix             As String
Private mvarLast_name2         As String
Private mvarFirst_name2        As String
Private mvarMiddle_name2       As String
Private mvarTitle              As String
Private mvarDegree             As String
Private mvarFirst_name_soundex As String
Private mvarLast_name_soundex  As String
Private mvarName_type_cid      As String
Private mvarUpdateFlag As String

Public Property Get UpdateFlag() As String:           UpdateFlag = mvarUpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mvarUpdateFlag = RHV:         End Property



Public Property Get Last_name() As String:                   Last_name = mvarLast_name:                   End Property
Public Property Let Last_name(ByVal RHV As String):          mvarLast_name = RHV:                         End Property

Public Property Get First_name() As String:                  First_name = mvarFirst_name:                 End Property
Public Property Let First_name(ByVal RHV As String):         mvarFirst_name = RHV:                        End Property

Public Property Get Middle_name() As String:                 Middle_name = mvarMiddle_name:               End Property
Public Property Let Middle_name(ByVal RHV As String):        mvarMiddle_name = RHV:                       End Property

Public Property Get Suffix() As String:                      Suffix = mvarSuffix:                         End Property
Public Property Let Suffix(ByVal RHV As String):             mvarSuffix = RHV:                            End Property

Public Property Get Last_name2() As String:                  Last_name2 = mvarLast_name2:                 End Property
Public Property Let Last_name2(ByVal RHV As String):         mvarLast_name2 = RHV:                        End Property

Public Property Get First_name2() As String:                 First_name2 = mvarFirst_name2:               End Property
Public Property Let First_name2(ByVal RHV As String):        mvarFirst_name2 = RHV:                       End Property

Public Property Get Middle_name2() As String:                Middle_name2 = mvarMiddle_name2:             End Property
Public Property Let Middle_name2(ByVal RHV As String):       mvarMiddle_name2 = RHV:                      End Property

Public Property Get Title() As String:                       Title = mvarTitle:                           End Property
Public Property Let Title(ByVal RHV As String):              mvarTitle = RHV:                             End Property

Public Property Get Degree() As String:                      Degree = mvarDegree:                         End Property
Public Property Let Degree(ByVal RHV As String):             mvarDegree = RHV:                            End Property

Public Property Get First_name_soundex() As String:          First_name_soundex = mvarFirst_name_soundex: End Property
Public Property Let First_name_soundex(ByVal RHV As String): mvarFirst_name_soundex = RHV:                End Property

Public Property Get Last_name_soundex() As String:           Last_name_soundex = mvarLast_name_soundex:   End Property
Public Property Let Last_name_soundex(ByVal RHV As String):  mvarLast_name_soundex = RHV:                 End Property

Public Property Get Name_type_cid() As String:               Name_type_cid = mvarName_type_cid:           End Property
Public Property Let Name_type_cid(ByVal RHV As String):      mvarName_type_cid = RHV:                     End Property


