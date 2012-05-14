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
Attribute VB_Name = "RequiredCommentsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RequiredCommentsType"
Attribute VB_Ext_KEY = "Member0" ,"RequiredCommentsType"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private mCol As Collection

Public Function Add(Prompt As String, Required As Boolean, Data_Type As hlRequiredCommentsDataTypes, _
            Default_Value As String, Help_Text As String, AnswerCollection As Collection, _
            RequiredCommentID As String, LowerLimit As Variant, UpperLimit As Variant, _
            MaxLength As Integer, Optional sKey As String) As RequiredCommentsType
            
            
    'create a new object
    Dim objNewMember As RequiredCommentsType
    Set objNewMember = New RequiredCommentsType


    'set the properties passed into the method
    objNewMember.Prompt = Prompt
    objNewMember.Required = Required
   
    objNewMember.Data_Type = Data_Type
    
    objNewMember.Default_Value = Default_Value
    objNewMember.Help_Text = Help_Text
    objNewMember.RequiredCommentID = RequiredCommentID
    objNewMember.LowerLimit = LowerLimit
    objNewMember.UpperLimit = UpperLimit
    objNewMember.MaxLength = MaxLength
    
    Set objNewMember.AnswerCollection = AnswerCollection
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function
Public Function AddUserResponse(RequiredCommentID As String, UserResponse As String) As RequiredCommentsType

    
    'create a new object
    Dim objNewMember As RequiredCommentsType
    Set objNewMember = New RequiredCommentsType
    
    objNewMember.RequiredCommentID = RequiredCommentID
    objNewMember.ReturnedAnswer = UserResponse
    mCol.Add objNewMember


    'return the object created
    Set AddUserResponse = objNewMember
    Set objNewMember = Nothing
End Function


Public Function AddViaHL7() As RequiredCommentsType
'NOTE: Only kept to maintain binary compatibility


'NOTE: HL7RetMsg  passed-on is really NOT optional, but written this way for binary compatibility
    
    Dim Prompt As String
    Dim bRequired As Boolean
    Dim Data_Type As hlRequiredCommentsDataTypes
    Dim Default_Value As String
    Dim Help_Text As String
    Dim AnswerCollection As Collection
    Dim RequiredCommentID As String
    Dim LowerLimit As Variant
    Dim UpperLimit As Variant
    Dim MaxLength As Integer
    Dim FileSource As hlSourceOfFileData
    Dim FileNumber As String
    Dim iLabZRRCount As Integer
    Dim iZRR As Integer
    Dim itemAnswer As Answers
    Dim iLabZRQCount As Integer
    Dim iZRQ As Integer
    Dim szTemp As String
    Dim oReq As RequiredCommentsType
   
    
    iLabZRQCount = iGetSegCount(gszHL7RetMsg, "ZRQ")
              
    For iZRQ = 1 To iLabZRQCount
        'create a new object
        Dim objNewMember As RequiredCommentsType
        Set objNewMember = New RequiredCommentsType
    
    
    
        Prompt = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.DisplayText)
        Data_Type = GetRequiredCommentsDataTypesEnumFromText(szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.ResponseType))
                                        
        szTemp = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.ResponseRequired)
        If szTemp = "Y" Then
            bRequired = True
        Else
            bRequired = False
        End If
        
        Default_Value = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.DefaultResponse)
        Help_Text = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.HelpText)
        LowerLimit = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.LowerLimit)
        UpperLimit = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.UpperLimit)
        RequiredCommentID = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.ReqComment)
        MaxLength = Val(szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.MaxLength))
        szTemp = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.SourceOfFileData)
        If szTemp = "1" Then
            FileSource = hlSourceCHCSI
        Else
            FileSource = hlSourceCHCSII
        End If
        FileNumber = szGetSegSeq(gszHL7RetMsg, iZRQ, "ZRQ", hlZRQSeg.FileNumber)
       
        
        'set the properties passed into the method
        objNewMember.Prompt = Prompt
        objNewMember.Required = bRequired
       
        objNewMember.Data_Type = Data_Type
        
        objNewMember.Default_Value = Default_Value
        objNewMember.Help_Text = Help_Text
        objNewMember.RequiredCommentID = RequiredCommentID
        objNewMember.LowerLimit = LowerLimit
        objNewMember.UpperLimit = UpperLimit
        objNewMember.MaxLength = MaxLength
        objNewMember.SourceOfFileData = FileSource
        objNewMember.FileNumber = FileNumber
        Set objNewMember.AnswerCollection = New Collection
        
      
        
        'Set objNewMember.AnswerCollection = Nothing
        
       ' If Len(sKey) = 0 Then
            mCol.Add objNewMember
       ' Else
       '     mCol.Add objNewMember, sKey
       ' End If
        'Set objNewMember = Nothing

    Next iZRQ
    
    'now add in the ZRRs
    iLabZRRCount = iGetSegCount(gszHL7RetMsg, "ZRR")
     
    For iZRR = 1 To iLabZRRCount
    '
    '
        RequiredCommentID = szGetSegSeq(gszHL7RetMsg, iZRR, "ZRR", hlZRRSeg.ReqComment)
        For Each oReq In mCol
            If oReq.RequiredCommentID = RequiredCommentID Then
                Set itemAnswer = New Answers
                If oReq.Data_Type = hlPointertoFile And oReq.SourceOfFileData = hlSourceCHCSI Then
                    itemAnswer.AnswerCode = szGetSegSeq(gszHL7RetMsg, iZRR, "ZRR", hlZRRSeg.CodeFile)
                    itemAnswer.AnswerText = szGetSegSeq(gszHL7RetMsg, iZRR, "ZRR", hlZRRSeg.ExpansionFile)
                Else
                    itemAnswer.AnswerCode = szGetSegSeq(gszHL7RetMsg, iZRR, "ZRR", hlZRRSeg.CodeSet)
                    itemAnswer.AnswerText = szGetSegSeq(gszHL7RetMsg, iZRR, "ZRR", hlZRRSeg.ExpansionSet)
                End If
        
                oReq.AnswerCollection.Add itemAnswer
                Exit For
            End If
        Next oReq
        
    Next iZRR
    'return the object created
    Set AddViaHL7 = objNewMember
    

End Function


Public Function Add_RequiredComments(iZRQ As Integer, HL7Msg As String) As RequiredCommentsType
    
    'Redesigned 1-08 to 1-11-2002 to manage multiple ZOCs; AAB (Taken from AddViaHL7)
    
Dim Prompt As String
Dim bRequired As Boolean
Dim Data_Type As hlRequiredCommentsDataTypes
Dim Default_Value As String
Dim Help_Text As String
        '' Dim AnswerCollection As Collection
Dim RequiredCommentID As String
Dim LowerLimit As Variant
Dim UpperLimit As Variant
Dim MaxLength As Integer
Dim FileSource As hlSourceOfFileData
Dim FileNumber As String
Dim iLabZRRCount As Integer
Dim iZRR As Integer
Dim itemAnswer As Answers
Dim iLabZRQCount As Integer
        ''Dim iZRQ As Integer
Dim szTemp As String
        ''Dim oReq As RequiredCommentsType
Dim oReqRequiredCommentID As String
    
            Dim objNewMember As RequiredCommentsType
            Set objNewMember = New RequiredCommentsType     'gets new AnswerCollection
                ''Set objNewMember.AnswerCollection = New Collection  ^^^^
                
                Prompt = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.DisplayText)
                Data_Type = GetRequiredCommentsDataTypesEnumFromText(szGetSegSeq(HL7Msg, _
                                                                iZRQ, "ZRQ", hlZRQSeg.ResponseType))
                                        
                szTemp = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.ResponseRequired)
                If szTemp = "Y" Then
                        bRequired = True
                Else
                        bRequired = False
                End If
                
                Default_Value = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.DefaultResponse)
                Help_Text = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.HelpText)
                LowerLimit = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.LowerLimit)
                UpperLimit = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.UpperLimit)
                RequiredCommentID = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.ReqComment)
                MaxLength = Val(szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.MaxLength))
                FileNumber = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.FileNumber)

                szTemp = szGetSegSeq(HL7Msg, iZRQ, "ZRQ", hlZRQSeg.SourceOfFileData)
                If szTemp = "1" Then
                        FileSource = hlSourceCHCSI
                Else
                        FileSource = hlSourceCHCSII
                End If
        
            'set the properties passed into the method
                objNewMember.Prompt = Prompt
                objNewMember.Required = bRequired
                objNewMember.Data_Type = Data_Type
                objNewMember.Default_Value = Default_Value
                objNewMember.Help_Text = Help_Text
                objNewMember.RequiredCommentID = RequiredCommentID
                objNewMember.LowerLimit = LowerLimit
                objNewMember.UpperLimit = UpperLimit
                objNewMember.MaxLength = MaxLength
                objNewMember.SourceOfFileData = FileSource
                objNewMember.FileNumber = FileNumber
                
                        ''  mCol.Add objNewMember
    
    'now add in the ZRRs ==========================================================================
    
                iLabZRRCount = iGetSegCount(HL7Msg, "ZRR")
                For iZRR = 1 To iLabZRRCount
                        oReqRequiredCommentID = szGetSegSeq(HL7Msg, iZRR, "ZRR", hlZRRSeg.ReqComment)
                        If (oReqRequiredCommentID = RequiredCommentID) And (Trim$(RequiredCommentID) <> "") Then
                                Set itemAnswer = New Answers
                                 If Data_Type = hlPointertoFile And FileSource = hlSourceCHCSI Then
                                        itemAnswer.AnswerCode = szGetSegSeq(HL7Msg, iZRR, "ZRR", hlZRRSeg.CodeFile)
                                        itemAnswer.AnswerText = szGetSegSeq(HL7Msg, iZRR, "ZRR", hlZRRSeg.ExpansionFile)
                                Else
                                        itemAnswer.AnswerCode = szGetSegSeq(HL7Msg, iZRR, "ZRR", hlZRRSeg.CodeSet)
                                        itemAnswer.AnswerText = szGetSegSeq(HL7Msg, iZRR, "ZRR", hlZRRSeg.ExpansionSet)
                                End If
                                objNewMember.AnswerCollection.Add itemAnswer
                                Set itemAnswer = Nothing
                        End If
                Next iZRR
                
                mCol.Add objNewMember
                
            'return the object created
                Set Add_RequiredComments = objNewMember
                Set objNewMember = Nothing
 
End Function







Public Property Get Item(vntIndexKey As Variant) As RequiredCommentsType
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

