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
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38135DF90135"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'##ModelId=38135DF90153
Private cAMSTypes As Collection
'##ModelId=38135DF9015D
Private cSurrogateIds As Collection

'##ModelId=38135DF90167
Public Property Get AMSTypes() As Collection
    Set AMSTypes = cAMSTypes
End Property

'##ModelId=38135DF90171
Public Property Let AMSTypes(Value As Collection)
    Set cAMSTypes = Value
End Property

'##ModelId=38135DF9017B
Public Property Get SurrogateIds() As Collection
    Set SurrogateIds = cSurrogateIds
End Property

'##ModelId=38135DF9018F
Public Property Let SurrogateIds(Value As Collection)
    Set cSurrogateIds = Value
End Property

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_InitProperties()
'
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
'
End Sub

Private Sub Class_Terminate()
'
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
'
End Sub

''Public Sub AddType(ByVal message_type_guid As String, ByVal Label As String, ByVal Priority As Integer, ByVal Interaction As String)
''
'''    Dim Connection As CHCSII_CONN.Conn
'''    Dim moGemsDAS As ICHCSII_DAS
'''
'''    On Error Resume Next
'''
'''    Set Connection = New CHCSII_CONN.Conn
'''    Set moGemsDAS = Connection.CHCSII_DAS
'''
'''    SQLOps.Execute "INSERT INTO alert_types (message_type_guid, label, priority, interaction) VALUES( '" & message_type_guid & "', '" & Label & "', " & Priority & ", '" & Interaction & "') WHERE NOT IN (SELECT message_type_guid FROM alert_types)"
'''
'''    Set moGemsDAS = nothing
''
''End Sub
''
''Public Sub Remove(ByVal message_type_guid As String)
''
'''    Dim Connection As CHCSII_CONN.Conn
'''    Dim moGemsDAS As ICHCSII_DAS
'''
'''    On Error Resume Next
'''
'''    Set Connection = New CHCSII_CONN.Conn
'''    Set moGemsDAS = Connection.CHCSII_DAS
'''
'''    SQLOps.Execute "DELETE FROM alert_types WHERE message_type_guid='" & message_type_guid & "'"
'''
'''    Set moGemsDAS = nothing
''
''End Sub
''
''
''Public Sub Update(ByVal message_type_guid As String, ByVal Label As String, ByVal Priority As Integer, ByVal Interaction As String)
''
'''    Dim Connection As CHCSII_CONN.Conn
'''    Dim moGemsDAS As ICHCSII_DAS
'''
'''    On Error Resume Next
'''
'''    Set Connection = New CHCSII_CONN.Conn
'''    Set moGemsDAS = Connection.CHCSII_DAS
'''
'''    SQLOps.Execute "UPDATE alert_types SET message_type_guid='" & message_type_guid & "', label='" & Label & "',priority=" & Priority & ", interaction='" & Interaction & "' WHERE message_type_guid='" & message_type_guid & "'"
'''
'''    Set moGemsDAS = nothing
''
''End Sub
''
