VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Workload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'''Workload Class to support new WAM project requirements for
'''the New Visit screen in the appointments module.

Private mvarAppt_Code            As String
Private mvarWorkload_Type        As String
Private mvarAppt_Duration        As String
Private mvarAppt_Description     As String
Public Property Let Appt_Description(ByVal vData As String)
    mvarAppt_Description = vData
End Property
Public Property Get Appt_Description() As String
    Appt_Description = mvarAppt_Description
End Property

Public Property Let Appt_Duration(ByVal vData As String)
    mvarAppt_Duration = vData
End Property

Public Property Get Appt_Duration() As String
    Appt_Duration = mvarAppt_Duration
End Property


Friend Function F_Deserialize(objFields As adodb.Fields) As Boolean
On Error Resume Next
   mvarAppt_Code = objFields("Appt_Code").Value & ""
   mvarAppt_Description = objFields("Description").Value & ""
   mvarAppt_Duration = objFields("Appt_Duration").Value & ""
   mvarWorkload_Type = objFields("Workload_Type").Value & ""

End Function

Public Property Let Workload_Type(ByVal vData As String)
    mvarWorkload_Type = vData
End Property

Public Property Get Workload_Type() As String
    Workload_Type = mvarWorkload_Type
End Property


Public Property Let Appt_Code(ByVal vData As String)
    mvarAppt_Code = vData
End Property

Public Property Get Appt_Code() As String
    Appt_Code = mvarAppt_Code
End Property



