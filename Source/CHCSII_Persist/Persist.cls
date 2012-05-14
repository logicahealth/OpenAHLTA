VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersistMethods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Persistence
Public Function Persist(ByRef obj As Object) As String
    Dim objII As tli.InterfaceInfo
    Dim objMI As tli.MemberInfo
    Dim objPersist As IPersist
    Dim objSR As tli.SearchResults
    Dim objSI As tli.SearchItem
    Dim j As Long
    
    Dim xml As DOMDocument30
    Dim e As MSXML2.IXMLDOMElement
    Dim n As MSXML2.IXMLDOMNode
    Dim b As MSXML2.IXMLDOMNode
    
    Set xml = New DOMDocument30
    Set objII = tli.InterfaceInfoFromObject(obj)
    
    'Get Strings, Longs, Etc
    Set e = xml.createElement(TypeName(obj))
    
    'Only Persist Read/Write Props
    Set objSR = objII.Members.GetFilteredMembers()
    For j = 1 To objSR.Count
        Set objSI = objSR.Item(j)
        If objSI.InvokeKinds = (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT) Then
            Select Case VarType(CallByName(obj, objSI.Name, VbGet))
            Case vbArray Or vbByte
                e.setAttribute objSI.Name, ""
                Set b = e.Attributes(e.Attributes.length - 1)
                b.dataType = "bin.hex"
                If Len(CStr(tli.InvokeHook(obj, objSI.Name, INVOKE_PROPERTYGET))) Then
                    b.nodeTypedValue = tli.InvokeHook(obj, objSI.Name, INVOKE_PROPERTYGET)
                End If
            Case vbObject
            Case Else 'Get Properties
                e.setAttribute objSI.Name, tli.InvokeHook(obj, objSI.Name, INVOKE_PROPERTYGET)
            End Select
        End If
    Next
    
    xml.appendChild e
    
    For Each objMI In objII.Members
        If objMI.InvokeKind = INVOKE_PROPERTYGET Then
        
            Select Case VarType(CallByName(obj, objMI.Name, VbGet))
            
            Case vbArray Or vbByte
            Case vbObject 'Get Sub Objects
            
                If TypeOf CallByName(obj, objMI.Name, VbGet) Is IPersist Then
                    Set objPersist = CallByName(obj, objMI.Name, VbGet)
                    
                    Set e = XMLElement(objPersist.Persist())
                    xml.documentElement.appendChild e
                    
                End If
                
            Case Else
            End Select
            
        End If
    Next
    
    
    'Debug.Assert TypeName(obj) <> "Layout"
    'xml.Save App.Path & "\test.xml"
    
    Persist = xml.xml

End Function
Public Function Restore(ByRef obj As Object, ByVal Persist As String) As Boolean
    
    Dim xml As MSXML2.DOMDocument30
    Dim bRet As Boolean
    Dim a As MSXML2.IXMLDOMAttribute
    Dim e As MSXML2.IXMLDOMElement
    Dim l As Long
    Dim objPersist As IPersist
    
    On Error GoTo errXIT:
    
    'Load XML String
    Set xml = New MSXML2.DOMDocument30
    
    bRet = xml.loadXML(Persist)
    If bRet = False Then
        Err.Raise vbObjectError
    End If
    
    'Set Simple Properties
    With xml.documentElement.Attributes
        For l = 0 To .length - 1
            Set a = .Item(l)
            
            Select Case VarType(CallByName(obj, a.Name, VbGet))
            Case vbArray Or vbByte
                a.dataType = "bin.hex"
                If IsNull(a.nodeTypedValue) Then
                    CallByName obj, a.Name, VbLet, vbNullString
                Else
                    CallByName obj, a.Name, VbLet, a.nodeTypedValue
                End If
            Case vbObject
            Case Else 'Get Properties
                CallByName obj, a.Name, VbLet, a.Value
            End Select
            
        Next
    End With
    
    'Build Collections
    With xml.documentElement.childNodes
        For l = 0 To .length - 1
            Set e = .Item(l)
            Set objPersist = CallByName(obj, e.nodeName, VbGet)
            If Not objPersist Is Nothing Then
                If objPersist.Restore(e.xml) <> True Then
                    Err.Raise vbObjectError
                End If
            End If
        Next
    End With
    
    Restore = True
    Exit Function
errXIT:

    Select Case Err.Number
    Case 451 'Property let procedure not defined and property get procedure did not return an object
        Resume Next 'Read Only Property
    Case 438 '"Object doesn't support this property or method", Property No Longer Exists
        Resume Next
    Case 13, 430 'Type mismatch; ie null string assigned to date
        Resume Next
    Case Else
        Restore = False
        Err.Raise Err.Number, "PersistMethods.Restore", "Error restoring " & TypeName(obj) & " with restoration string: '" & Persist & "'", vbExclamation, App.Title
    End Select
    
End Function

Public Function PersistCollection(ByRef obj As Object, ByRef col As Collection) As String
    
    Dim xml As DOMDocument30
    Dim e As MSXML2.IXMLDOMElement
    Dim objPersist As IPersist
    
    Set xml = New DOMDocument30
    Set e = xml.createElement(TypeName(obj))
    xml.appendChild e
    
    For Each objPersist In col
        Set e = XMLElement(objPersist.Persist())
        xml.documentElement.appendChild e
    Next
    
    PersistCollection = xml.xml
        
End Function
Private Function XMLElement(ByVal sz As String) As IXMLDOMElement

    Dim xml As DOMDocument30
    
    Set xml = New DOMDocument30
    xml.loadXML sz
    Set XMLElement = xml.documentElement
    
End Function

Public Function RestoreCollection(ByRef objParent As IPersist, ByRef col As Collection, _
                                  ByVal Persist As String) As Boolean

    
    Dim xml As MSXML2.DOMDocument30
    Dim bRet As Boolean
    Dim a As MSXML2.IXMLDOMAttribute
    Dim e As MSXML2.IXMLDOMElement
    Dim l As Long
    Dim objPersist As IPersist
    Dim obj As Object
    
    On Error GoTo errXIT:
    
    'Load XML String
    Set xml = New MSXML2.DOMDocument30
    bRet = xml.loadXML(Persist)
    If bRet = False Then
        Err.Raise vbObjectError
    End If
    
    'Clear Collection
    Set col = New Collection
    
    'Create Objects, Fill Them In, Add To Collection
    With xml.documentElement.childNodes
        For l = 0 To .length - 1
            Set e = .Item(l)
            'Have Parent Generate Our new objects
            Set obj = objParent.ChildFactory()
            Set objPersist = obj
            If objPersist.Restore(e.xml) Then
                col.Add obj
            Else
                Err.Raise vbObjectError
            End If
        Next
    End With
   
    RestoreCollection = True
    Exit Function
    
errXIT:
    RestoreCollection = False
    MsgBox "Error restoring collection:" & TypeName(objParent) & " with restoration string: '" & Persist & "'", vbExclamation, App.Title
    
End Function


Public Function Describe(ByRef obj As Object) As String

    Dim objII As tli.InterfaceInfo
    Dim objMI As tli.MemberInfo
    Dim objPersist As IPersist
    Dim objSR As tli.SearchResults
    Dim objSI As tli.SearchItem
    Dim objMs As tli.Members
    Dim szXML As String
    Dim j As Long
    
    Dim xml As DOMDocument30
    Dim e As MSXML2.IXMLDOMElement
    Dim e2 As MSXML2.IXMLDOMElement
    Dim e3 As MSXML2.IXMLDOMElement
    Dim n As MSXML2.IXMLDOMNode
    Dim b As MSXML2.IXMLDOMNode
    
    Set xml = New DOMDocument30
    Set objII = tli.InterfaceInfoFromObject(obj)
    
    'Get Strings, Longs, Etc
    Set e = xml.createElement(TypeName(obj))
    
    'Describe Properties
    Set objSR = objII.Members.GetFilteredMembers()
    e.setAttribute "HelpString", objII.HelpString
    For j = 1 To objSR.Count
        
        Set objSI = objSR.Item(j)
        Set objMI = objII.GetMember(objSI.MemberId)
             
        szXML = DescribeMember(objMI)
        If Len(szXML) Then
            Set e2 = XMLElement(szXML)
            e.appendChild e2
        End If

    Next
   
    xml.appendChild e
    

    
    Describe = xml.xml

End Function



Private Function DescribeMember(ByRef objMI As tli.MemberInfo) As String

    Dim objII As tli.InterfaceInfo
    Dim objPersist As IPersist
    Dim objSR As tli.SearchResults
    Dim objSI As tli.SearchItem
    Dim objMs As tli.Members
    Dim szXML As String
    Dim j As Long
    Dim szParms As String
    Dim objParm As ParameterInfo
    
    Dim xml As DOMDocument30
    Dim e As MSXML2.IXMLDOMElement
    Dim e2 As MSXML2.IXMLDOMElement
    Dim e3 As MSXML2.IXMLDOMElement
    Dim n As MSXML2.IXMLDOMNode
    Dim b As MSXML2.IXMLDOMNode
    Dim objMI2 As MemberInfo
    
    Set xml = New DOMDocument30
    'Create member Info
     If objMI.ReturnType.TypeInfo Is Nothing Then
        Set e = xml.createElement("Property")
        e.setAttribute "Name", objMI.Name
        
        Select Case objMI.ReturnType.VarType
        Case TliVarType.VT_BOOL
            e.setAttribute "Type", "BOOL"
        Case TliVarType.VT_DATE
            e.setAttribute "Type", "DATE"
        Case TliVarType.VT_VARIANT
            e.setAttribute "Type", "VARIANT"
        Case TliVarType.VT_ARRAY
            e.setAttribute "Type", "ARRAY"
        Case TliVarType.VT_DECIMAL
            e.setAttribute "Type", "DECIMAL"
        Case TliVarType.VT_INT
            e.setAttribute "Type", "INT"
        Case TliVarType.VT_LPSTR
            e.setAttribute "Type", "LPSTR"
        Case TliVarType.VT_I4
            e.setAttribute "Type", "LONG"
        Case TliVarType.VT_BSTR
            e.setAttribute "Type", "STRING"
        Case vbArray Or vbByte
            e.setAttribute "Type", "Byte Array"
        Case TliVarType.VT_VOID
            Set e = xml.createElement("Method")
            e.setAttribute "Name", objMI.Name
        Case TliVarType.VT_EMPTY

        Case Else
            Stop
        End Select
    
        e.setAttribute "HelpString", objMI.HelpString
        
    ElseIf objMI.ReturnType.TypeInfo.TypeKind = TKIND_ENUM Then
        Set e = xml.createElement("Property")
        e.setAttribute "Name", objMI.Name
        e.setAttribute "Returns", objMI.ReturnType.TypeInfo.Name
        e.setAttribute "HelpString", objMI.HelpString
    Else
    'Methods Append
        Set e = xml.createElement("Method")
        e.setAttribute "Name", objMI.Name
        e.setAttribute "Returns", objMI.ReturnType.TypeInfo.Name

        If objMI.Parameters.Count Then
            For Each objParm In objMI.Parameters
                szParms = szParms & ", " & objParm.Name
            Next
            e.setAttribute "Parms", Right(szParms, Len(szParms) - 2)
            
        End If
        
        If objMI.ReturnType.IsExternalType = False Then
            Set objSR = objMI.ReturnType.TypeInfo.Members.GetFilteredMembers
            
            Set e2 = xml.createElement(objMI.ReturnType.TypeInfo.Name)
            For j = 1 To objSR.Count
                Set objSI = objSR.Item(j)
                Set objMI2 = objMI.ReturnType.TypeInfo.GetMember(objSI.MemberId)
                
                szXML = DescribeMember(objMI2)
                If Len(szXML) Then
                    Set e3 = XMLElement(szXML)
                    e2.appendChild e3
                End If
            Next
            
            e.setAttribute "HelpString", objMI.ReturnType.TypeInfo.HelpString
            e.appendChild e2
        End If
    End If
   
   
'
'    'Objects Append
'    For j = 1 To objSR.Count
'        Set objSI = objSR.Item(j)
'        Set objMI = objII.GetMember(objSI.MemberId)
'
'        If Not objMI.ReturnType.TypeInfo Is Nothing Then
'            If objMI.Parameters.Count = 0 Then
'                If objMI.ReturnType.TypeInfo.TypeKind <> TKIND_ENUM Then
'                Set e2 = xml.createElement("Method")
'                e2.setAttribute "Name", objMI.Name
'                e2.setAttribute "Object", objMI.ReturnType.TypeInfo.Name
'                e2.setAttribute "HelpString", objMI.HelpString
'
'                szXML = Describe(CallByName(obj, objMI.Name, VbGet))
'                Set e3 = XMLElement(szXML)
'                e2.appendChild e3
'                e.appendChild e2
'                End If
'            End If
'        End If
'
'    Next
    
'    xml.appendChild e
    

    
    DescribeMember = e.xml

End Function
Function DescribeObject(ByRef objMI As MemberInfo)

    Dim objII As tli.InterfaceInfo
    Dim objPersist As IPersist
    Dim objSR As tli.SearchResults
    Dim objSI As tli.SearchItem
    Dim objMs As tli.Members
    Dim szXML As String
    Dim j As Long
    Dim szParms As String
    Dim objParm As ParameterInfo
    
    Dim xml As DOMDocument30
    Dim e As MSXML2.IXMLDOMElement
    Dim e2 As MSXML2.IXMLDOMElement
    Dim e3 As MSXML2.IXMLDOMElement
    Dim n As MSXML2.IXMLDOMNode
    Dim b As MSXML2.IXMLDOMNode
    
    Set e = XMLElement(DescribeMember(objMI))
    
    
End Function

