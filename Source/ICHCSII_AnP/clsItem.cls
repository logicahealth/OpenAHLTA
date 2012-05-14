VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Class Name:    clsItem.cls
'Author:        Phillip Crowder
'Created:       9/19/01
'Description:   This is a single model for orders and diagnosis related items.  It is a
'               generic property class.  A dictionary is used to store the property values.
'Revisions:

Option Explicit

'An order or related item can be one of these types.  If the type is undefined then
'the developer has made an error.
Public Enum enumItemType
  itUndefined
  itProcedure
  itConsult
  itRad
  itLab
  itMed
  itPatientInstr
  itInPatient 'SCR 56524 KMonroe 11/03/04
End Enum

'Possible states for orders and related items.  If the state is unknown then the
'developer has made an error.  New, Delete, Remove, Saved, Queued, New Problem and Modified are
'valid states for orders.  Associate, Associated, Unassociate, Unassociated, Queued, and
'New Problem are vaild states for related items.
'New          - Order added to AnP but not saved
'Delete       - Order is flagged for deletion but has not been deleted
'Remove       - Order has been deleted
'Saved        - Order has been saved
'Associate    - Order has been associated to a diagnosis but the relationship has not been saved
'Associated   - Association of an order to a diagnosis has been saved
'Unassociate  - The relationship between an order and a diagnosis has been broken in AnP but not removed
'Unassociated - Association of an order to a diagnosis has been removed
'Queued       - Order has been submitted to AnP but not processed by the OE module.
'               Once queued items are submitted to OE their states are changed to New for Orders and
'               Assoicate for Related Items.
'Modified     - Order has been modified but the changes have not been saved.
'               Rad, Lab and Med orders can be modified in the OE module.
Public Enum enumItemState
  isUnknown
  isNew
  isDelete
  isRemove
  isSaved
  isAssociate
  isAssociated
  isUnassociate
  isUnassociated
  isQueued
  isModified
End Enum

Private mintState As enumItemState      'Current state of order or related item
Private mintType As enumItemType        'Current type of order or related item

Private mstrKey As String               'Used to link an order to related items.
                                        'Orders store more detailed information then
                                        'related items so this key is used to get
                                        'information from the order for a related item.
                                        'Orders do not have a unique key till they are
                                        'saved.

Private mdicProps As Dictionary         'Collection of property values such as SnoID, DataId, etc.

Private Sub Class_Initialize()
  Set mdicProps = New Dictionary
End Sub

Private Sub Class_Terminate()
  Set mdicProps = Nothing
End Sub

'This property is required for backwards compatability with encounter.
Public Property Get DataId() As String
  DataId = mdicProps.Item("DataId")
End Property

Public Property Let ItemState(ByVal NewValue As enumItemState)
  mintState = NewValue
End Property

Public Property Get ItemState() As enumItemState
  ItemState = mintState
End Property

Public Property Let ItemType(ByVal NewValue As enumItemType)
  mintType = NewValue
End Property

Public Property Get ItemType() As enumItemType
  ItemType = mintType
End Property

Public Property Let Key(ByVal NewValue As String)
  If Len(Trim$(mstrKey)) = 0 Then mstrKey = NewValue
End Property

Public Property Get Key() As String
  Key = mstrKey
End Property

Public Property Get PropExists(ByVal PropName As String) As Boolean
  PropExists = mdicProps.Exists(PropName)
End Property

Public Property Let PropValue(ByVal PropName As String, ByVal NewValue As Variant)
  mdicProps.Item(PropName) = NewValue
End Property

Public Property Get PropValue(ByVal PropName As String) As Variant
  If mdicProps.Exists(PropName) Then PropValue = mdicProps.Item(PropName)
End Property

