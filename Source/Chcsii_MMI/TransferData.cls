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
Attribute VB_Name = "TransferData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'To delete, set the collection member type to 0. WriteUpdate to commit the changes.
'To add, move to Eof and set the Object to the new member Type value. WriteUpdate to
' commit the changes.
Public Sub CDRGenericId(FromId As Member, ToId As Member)
    Dim bFound As Boolean
    'FromId.GenericID As GenericId, ToId.GenericID As GenericId)
    
    If FromId.GenericId Is Nothing Then Exit Sub
    
    FromId.GenericId.GotoFirst
    
    If ToId.GenericId Is Nothing Then
        Do While Not FromId.GenericId.EOF
            ToId.GenericId = FromId.GenericId.Type
            With ToId.GenericId
                .AssigningAuthority = FromId.GenericId.AssigningAuthority
                .Value = FromId.GenericId.Value
            End With
            FromId.GenericId.GotoNext
            ToId.GenericId.GotoNext
        Loop
    Else
        Do While Not FromId.GenericId.EOF
            ToId.GenericId.GotoFirst
            bFound = False
            Do While Not ToId.GenericId.EOF
                If ToId.GenericId.Type = FromId.GenericId.Type Then
                    With ToId.GenericId
                        .AssigningAuthority = FromId.GenericId.AssigningAuthority
                        .Value = FromId.GenericId.Value
                    End With
                    bFound = True
                    Exit Do
                End If
                ToId.GenericId.GotoNext
            Loop
            If Not bFound Then
                ToId.GenericId.Type = FromId.GenericId.Type
                With ToId.GenericId
                    .AssigningAuthority = FromId.GenericId.AssigningAuthority
                    .Value = FromId.GenericId.Value
                End With
            End If
            FromId.GenericId.GotoNext
        Loop
    End If
        
End Sub
Public Sub CDRGenericData(FromData As Member, ToData As Member)
    Dim bFound As Boolean
    'FromData.GenericData As GenericData, ToData.GenericData As GenericData
    If FromData.GenericData Is Nothing Then Exit Sub
    
    FromData.GenericData.GotoFirst
    
    If ToData.GenericData Is Nothing Then
        Do While Not FromData.GenericData.EOF
            ToData.GenericData = FromData.GenericData.Type
            With ToData.GenericData
                .ErrorText = FromData.GenericData.ErrorText
                .ValueNcid = FromData.GenericData.ValueNcid
                .ValueRsid = FromData.GenericData.ValueRsid
                .ValueText = FromData.GenericData.ValueText
            End With
            FromData.GenericData.GotoNext
            ToData.GenericData.GotoNext
        Loop
    Else
        Do While Not FromData.GenericData.EOF
            ToData.GenericData.GotoFirst
            bFound = False
            Do While Not ToData.GenericData.EOF
                If ToData.GenericData.Type = FromData.GenericData.Type Then
                    With ToData.GenericData
                        .ErrorText = FromData.GenericData.ErrorText
                        .ValueNcid = FromData.GenericData.ValueNcid
                        .ValueRsid = FromData.GenericData.ValueRsid
                        .ValueText = FromData.GenericData.ValueText

                    End With
                    bFound = True
                    Exit Do
                End If
                ToData.GenericData.GotoNext
            Loop
            If Not bFound Then
                ToData.GenericData.Type = FromData.GenericData.Type
                With ToData.GenericData
                    .ErrorText = FromData.GenericData.ErrorText
                    .ValueNcid = FromData.GenericData.ValueNcid
                    .ValueRsid = FromData.GenericData.ValueRsid
                    .ValueText = FromData.GenericData.ValueText
                End With
            End If
            FromData.GenericData.GotoNext
        Loop
    End If
        
End Sub

Public Sub CDRAddress(FromAddress As Member, ToAddress As Member)
    Dim bFound As Boolean
    'FromAddress.Address As Address, ToAddress.Address As Address
    If FromAddress.Address Is Nothing Then Exit Sub
    
    FromAddress.Address.GotoFirst
    
    If ToAddress.Address Is Nothing Then
        Do While Not FromAddress.Address.EOF
            ToAddress.Address = FromAddress.Address.Type
            With ToAddress.Address
                .City = FromAddress.Address.City
                .Country = FromAddress.Address.Country
                .County = FromAddress.Address.County
                .PostCode = FromAddress.Address.PostCode
                .State = FromAddress.Address.State
                .Street = FromAddress.Address.Street
                .Street2 = FromAddress.Address.Street2
            End With
            FromAddress.Address.GotoNext
            ToAddress.Address.GotoNext
        Loop
    Else
        Do While Not FromAddress.Address.EOF
            ToAddress.Address.GotoFirst
            bFound = False
            Do While Not ToAddress.Address.EOF
                If ToAddress.Address.Type = FromAddress.Address.Type Then
                    With ToAddress.Address
                        .City = FromAddress.Address.City
                        .Country = FromAddress.Address.Country
                        .County = FromAddress.Address.County
                        .PostCode = FromAddress.Address.PostCode
                        .State = FromAddress.Address.State
                        .Street = FromAddress.Address.Street
                        .Street2 = FromAddress.Address.Street2
                    End With
                    bFound = True
                    Exit Do
                End If
                ToAddress.Address.GotoNext
            Loop
            If Not bFound Then
                ToAddress.Address.Type = FromAddress.Address.Type
                With ToAddress.Address
                    .City = FromAddress.Address.City
                    .Country = FromAddress.Address.Country
                    .County = FromAddress.Address.County
                    .PostCode = FromAddress.Address.PostCode
                    .State = FromAddress.Address.State
                    .Street = FromAddress.Address.Street
                    .Street2 = FromAddress.Address.Street2
                End With
            End If
            FromAddress.Address.GotoNext
        Loop
    End If
    
End Sub
Public Sub CDRFacility(FromFacility As Member, ToFacility As Member)
    Dim bFound As Boolean
    'FromFacility.Facility As Facility, ToFacility.Facility As Facility
    If FromFacility.Facility Is Nothing Then Exit Sub
    
    FromFacility.Facility.GotoFirst
    
    If ToFacility.Facility Is Nothing Then
        Do While Not FromFacility.Facility.EOF
            ToFacility.Facility = FromFacility.Facility.Type
            With ToFacility.Facility
                .MedRecNum = FromFacility.Facility.MedRecNum
                .Name = FromFacility.Facility.Name
                .RadiologyNum = FromFacility.Facility.RadiologyNum
            End With
            FromFacility.Facility.GotoNext
            ToFacility.Facility.GotoNext
        Loop
    Else
        Do While Not FromFacility.Facility.EOF
            ToFacility.Facility.GotoFirst
            bFound = False
            Do While Not ToFacility.Facility.EOF
                If ToFacility.Facility.Type = FromFacility.Facility.Type Then
                    With ToFacility.Facility
                        .MedRecNum = FromFacility.Facility.MedRecNum
                        .Name = FromFacility.Facility.Name
                        .RadiologyNum = FromFacility.Facility.RadiologyNum
                    End With
                    bFound = True
                    Exit Do
                End If
                ToFacility.Facility.GotoNext
            Loop
            If Not bFound Then
                ToFacility.Facility.Type = FromFacility.Facility.Type
                With ToFacility.Facility
                    .MedRecNum = FromFacility.Facility.MedRecNum
                    .Name = FromFacility.Facility.Name
                    .RadiologyNum = FromFacility.Facility.RadiologyNum
                End With
            End If
            FromFacility.Facility.GotoNext
        Loop
    End If
End Sub
Public Sub CDRPhone(FromPhone As Member, ToPhone As Member)
    Dim bFound As Boolean
    'FromPhone.Phone As Phone, ToPhone.Phone As Phone
    If FromPhone.Phone Is Nothing Then Exit Sub
    
    FromPhone.Phone.GotoFirst
    
    If ToPhone.Phone Is Nothing Then
        Do While Not FromPhone.Phone.EOF
            ToPhone.Phone = FromPhone.Phone.Type
            With ToPhone.Phone
                .AreaCode = FromPhone.Phone.AreaCode
                .CountryCode = FromPhone.Phone.CountryCode
                .Extension = FromPhone.Phone.Extension
                .LocalNumber = FromPhone.Phone.LocalNumber
            End With
            FromPhone.Phone.GotoNext
            ToPhone.Phone.GotoNext
        Loop
    Else
        Do While Not FromPhone.Phone.EOF
            ToPhone.Phone.GotoFirst
            bFound = False
            Do While Not ToPhone.Phone.EOF
                If ToPhone.Phone.Type = FromPhone.Phone.Type Then
                    With ToPhone.Phone
                        .AreaCode = FromPhone.Phone.AreaCode
                        .CountryCode = FromPhone.Phone.CountryCode
                        .Extension = FromPhone.Phone.Extension
                        .LocalNumber = FromPhone.Phone.LocalNumber
                    End With
                    bFound = True
                    Exit Do
                End If
                ToPhone.Phone.GotoNext
            Loop
            If Not bFound Then
                ToPhone.Phone.Type = FromPhone.Phone.Type
                With ToPhone.Phone
                    .AreaCode = FromPhone.Phone.AreaCode
                    .CountryCode = FromPhone.Phone.CountryCode
                    .Extension = FromPhone.Phone.Extension
                    .LocalNumber = FromPhone.Phone.LocalNumber
                End With
            End If
            FromPhone.Phone.GotoNext
        Loop
    End If
    
End Sub
Public Sub CDRName(FromName As Member, ToName As Member)
    Dim bFound As Boolean
    'FromName.Name As Name, ToName.Name As Name
    If FromName.Name Is Nothing Then Exit Sub
    
    FromName.Name.GotoFirst
    
    If ToName.Name Is Nothing Then
        Do While Not FromName.Name.EOF
            ToName.Name = FromName.Name.Type
            With ToName.Name
                .Degree = FromName.Name.Degree
                .First = FromName.Name.First
                .Full = FromName.Name.Full
                .Last = FromName.Name.Last
                .Last2 = FromName.Name.Last2
                .Middle = FromName.Name.Middle
                .Middle2 = FromName.Name.Middle2
                .Suffix = FromName.Name.Suffix
                .Title = FromName.Name.Title

            End With
            FromName.Name.GotoNext
            ToName.Name.GotoNext
        Loop
    Else
        Do While Not FromName.Name.EOF
            ToName.Name.GotoFirst
            bFound = False
            Do While Not ToName.Name.EOF
                If ToName.Name.Type = FromName.Name.Type Then
                    With ToName.Name
                        .Degree = FromName.Name.Degree
                        .First = FromName.Name.First
                        .Full = FromName.Name.Full
                        .Last = FromName.Name.Last
                        .Last2 = FromName.Name.Last2
                        .Middle = FromName.Name.Middle
                        .Middle2 = FromName.Name.Middle2
                        .Suffix = FromName.Name.Suffix
                        .Title = FromName.Name.Title

                    End With
                    bFound = True
                    Exit Do
                End If
                ToName.Name.GotoNext
            Loop
            If Not bFound Then
                ToName.Name.Type = FromName.Name.Type
                With ToName.Name
                    .Degree = FromName.Name.Degree
                    .First = FromName.Name.First
                    .Full = FromName.Name.Full
                    .Last = FromName.Name.Last
                    .Last2 = FromName.Name.Last2
                    .Middle = FromName.Name.Middle
                    .Middle2 = FromName.Name.Middle2
                    .Suffix = FromName.Name.Suffix
                    .Title = FromName.Name.Title
                End With
            End If
            FromName.Name.GotoNext
        Loop
    End If
    
End Sub

Public Function TransferFromRSobj(rsobj As Collection) As MmiDoc
    Dim rs          As Recordset
    Dim sTable      As String
    Dim oDoc        As MmiDoc
    Dim oMember     As Member
    Dim vObj        As Variant   'v.Tablename, v.DBRecordSet
    
    On Error GoTo ErrHandler
    
    Set oDoc = New thr_MmiDoc
    Call oDoc.CreateMember("patient")
    Set oMember = oDoc.Member

    For Each vObj In rsobj
        If Trim(UCase(vObj.TableName)) = "MMI_ID" Then
            Set rs = vObj.DBRecordSet
            
            If rs.RecordCount Then
                rs.MoveFirst
                oMember.Name = "187" 'ncid_current_name
                
                'oMember.Name.Last = rs!Last
                oMember.SSN = rs!SSN
                Exit For
            End If
        End If
    Next vObj
    
    If oDoc.DoSearch(oDoc) Then
        Set oMember = oDoc.Member
    Else 'new patient apparently
        
    End If
    
    'the easy way out
    On Error Resume Next

    With oMember
        .Adopted = rs!Adoption_flag
        .BirthDate = rs!BIRTH_DATE
        '.Confidence = rs!Confidence
        .DateOfDeath = IIf(IsNull(rs!DEATH_DATE), vbNull, rs!DEATH_DATE)
        '.DOBEstimated = rs!DOB_EST_FLAG
        .Expired = rs!EXPIRED_FLAG
        .GovtHealthNumber = rs!GOVT_HEALTH_NUMBER
        .Guarantor = rs!GUARANTOR_FLAG
        .Institution = rs!INSTITUTIONAL_FLAG
        .MaidenName = rs!MAIDEN_NAME
        .MaritalStatus = rs!MARITAL_STATUS_CID
        .MomMaidenName = IIf(IsNull(rs!MOTHER_NAME), "", rs!MOTHER_NAME)
        .OrganDonor = rs!ORGAN_DONOR_FLAG
        .Passport = rs!PASSPORT_NUMBER
        .Patient = rs!PATIENT_FLAG
        .Race = rs!RACE_CID
        .ReferenceId = rs!REFERENCE
        .Religion = rs!RELIGION_CID
        '.Self = rs!Self
        .Sex = rs!Sex
        .SSN = rs!SSN
        .SSNCorrect = rs!CORRECT_SSN
        .Unidentified = rs!UNIDENTIFIED_FLAG
        '.UnitNumber = rs!UNIT_NUMBER
    End With

    Dim bFound As Boolean

    For Each vObj In rsobj
        Set rs = vObj.DBRecordSet
        sTable = Trim(UCase(vObj.TableName))

        Select Case sTable
            Case "MMI_GENERIC_ID"
                
                If rs Is Nothing Then GoTo NextRec
                
                rs.MoveFirst
                
                If oMember.GenericId Is Nothing Then
                    Do While Not rs.EOF
                        oMember.GenericId = rs!ID_TYPE_NCID
                        With oMember.GenericId
                            .AssigningAuthority = IIf(IsNull(rs!ASSIGNING_AUTHORITY_NCID), 0, rs!ASSIGNING_AUTHORITY_NCID)
                            .Value = rs!ID_VALUE
                        End With
                        rs.MoveNext
                        oMember.GenericId.GotoNext
                    Loop
                Else
                    Do While Not rs.EOF
                        oMember.GenericId.GotoFirst
                        bFound = False
                        Do While Not oMember.GenericId.EOF
                            If Val(oMember.GenericId.Type) = rs!ID_TYPE_NCID Then
                                With oMember.GenericId
                                    .AssigningAuthority = IIf(IsNull(rs!ASSIGNING_AUTHORITY_NCID), "0", rs!ASSIGNING_AUTHORITY_NCID)
                                    .Value = rs!ID_VALUE
                                End With
                                bFound = True
                                Exit Do
                            End If
                            oMember.GenericId.GotoNext
                        Loop
                        If Not bFound Then
                            oMember.GenericId.Type = rs!ID_TYPE_NCID
                            With oMember.GenericId
                                .AssigningAuthority = IIf(IsNull(rs!ASSIGNING_AUTHORITY_NCID), vbNull, rs!ASSIGNING_AUTHORITY_NCID)
                                .Value = rs!ID_VALUE
                            End With
                        End If
'                        oMember.GenericId.GotoNext
                        rs.MoveNext
                    Loop
                End If

            Case "MMI_GENERICDATA"
                If rs Is Nothing Then GoTo NextRec
                
                rs.MoveFirst
                
                If oMember.GenericData Is Nothing Then
                    Do While Not rs.EOF
                        oMember.GenericData = rs!Type_NCID
                        With oMember.GenericData
                            .ErrorText = rs!Error_Text
                            .ValueNcid = rs!Value_NCID
                            .ValueRsid = rs!Value_Rsid
                            .ValueText = rs!Value_Text
                        End With
                        rs.MoveNext
                        oMember.GenericData.GotoNext
                    Loop
                Else
                    Do While Not rs.EOF
                        oMember.GenericData.GotoFirst
                        bFound = False
                        Do While Not oMember.GenericData.EOF
                            If Val(oMember.GenericData.Type) = rs!Type_NCID Then
                                If (rs!Type_NCID <> 204685 And rs!Type_NCID <> 204687) _
                                   Or ((rs!Type_NCID = 204685 Or rs!Type_NCID = 204687) And _
                                       Val(oMember.GenericData.ValueNcid) = rs!Value_NCID) Then
                                    With oMember.GenericData
                                        .ErrorText = rs!Error_Text
                                        .ValueNcid = rs!Value_NCID
                                        .ValueRsid = rs!Value_Rsid
                                        .ValueText = rs!Value_Text
                
                                    End With
                                    bFound = True
                                    Exit Do
                                End If
                            End If
                            oMember.GenericData.GotoNext
                        Loop
                        If Not bFound Then
                            oMember.GenericData.Type = rs!Type_NCID
                            With oMember.GenericData
                                .ErrorText = rs!Error_Text
                                .ValueNcid = rs!Value_NCID
                                .ValueRsid = rs!Value_Rsid
                                .ValueText = rs!Value_Text
                            End With
                            
                        End If
                        rs.MoveNext
                    Loop
                End If
            Case "MMI_ADDRESS"
                 If rs Is Nothing Then GoTo NextRec
                
                rs.MoveFirst
                
                If oMember.Address Is Nothing Then
                    Do While Not rs.EOF
                        oMember.Address = rs!ADDRESS_TYPE_CID
                        With oMember.Address
                            .City = rs!City
                            .Country = rs!Country
                            .County = rs!County
                            .PostCode = rs!POST_CODE
                            .State = rs!State
                            .Street = rs!Street_1
                            .Street2 = rs!Street_2
                        End With
                        rs.MoveNext
                        oMember.Address.GotoNext
                    Loop
                Else
                    Do While Not rs.EOF
                        oMember.Address.GotoFirst
                        bFound = False
                        Do While Not oMember.Address.EOF
                            If Val(oMember.Address.Type) = rs!ADDRESS_TYPE_CID Then
                                With oMember.Address
                                    .City = rs!City
                                    .Country = rs!Country
                                    .County = rs!County
                                    .PostCode = rs!POST_CODE
                                    .State = rs!State
                                    .Street = rs!Street_1
                                    .Street2 = rs!Street_2
                                End With
                                bFound = True
                                Exit Do
                            End If
                            oMember.Address.GotoNext
                        Loop
                        If Not bFound Then
                            oMember.Address.Type = rs!ADDRESS_TYPE_CID
                            With oMember.Address
                                .City = rs!City
                                .Country = rs!Country
                                .County = rs!County
                                .PostCode = rs!POST_CODE
                                .State = rs!State
                                .Street = rs!Street_1
                                .Street2 = rs!Street_2
                            End With
                        End If
                        rs.MoveNext
                    Loop
                End If
                                                
            Case "MMI_NAME"
                If rs Is Nothing Then GoTo NextRec
                
                rs.MoveFirst
                
                If oMember.Name Is Nothing Then
                    Do While Not rs.EOF
                        oMember.Name = rs!Type
                        With oMember.Name
                            .Degree = rs!Degree
                            .First = rs!First
                            .Full = rs!Full
                            .Last = rs!Last
                            .Last2 = rs!Last2
                            .Middle = rs!Middle
                            .Middle2 = rs!Middle2
                            .Suffix = rs!Suffix
                            .Title = rs!Title
            
                        End With
                        rs.MoveNext
                        oMember.Name.GotoNext
                    Loop
                Else
                    Do While Not rs.EOF
                        oMember.Name.GotoFirst
                        bFound = False
                        Do While Not oMember.Name.EOF
                            'eic, has name data but doesn't know the type, default to NCID_CURRENT_NAME
                            If oMember.Name.Type = IIf(IsNull(rs!Name_type_cid), "187", rs!Name_type_cid) Then
                                With oMember.Name
                                    .Degree = rs!Degree
                                    .First = rs!First_name
                                    '.Full = rs!Full
                                    .Last = rs!Last_name
                                    .Last2 = rs!Last_name_2
                                    .Middle = rs!Middle_name
                                    .Middle2 = rs!Middle_name_2
                                    .Suffix = rs!Suffix
                                    .Title = rs!Title
                                End With
                                bFound = True
                                Exit Do
                            End If
                            oMember.Name.GotoNext
                        Loop
                        If Not bFound Then
                            oMember.Name.Type = IIf(IsNull(rs!Name_type_cid), "187", rs!Name_type_cid)
                            With oMember.Name
                                .Degree = rs!Degree
                                .First = rs!First_name
                                '.Full = rs!Full
                                .Last = rs!Last_name
                                .Last2 = rs!Last_name_2
                                .Middle = rs!Middle_name
                                .Middle2 = rs!Middle_name_2
                                .Suffix = rs!Suffix
                                .Title = rs!Title
                            End With
                        End If
                        rs.MoveNext
                    Loop
                End If
            
            Case "MMI_PHONE"
                If rs Is Nothing Then GoTo NextRec
                
                rs.MoveFirst
                
                If oMember.Phone Is Nothing Then
                    Do While Not rs.EOF
                        oMember.Phone = rs!PHONE_TYPE_CID
                        With oMember.Phone
                            .AreaCode = IIf(IsNull(rs!AREA_CODE), "", rs!AREA_CODE)
                            .CountryCode = IIf(IsNull(rs!COUNTRY_CODE), "", rs!COUNTRY_CODE)
                            .Extension = IIf(IsNull(rs!INTERNAL_EXTENSION), "", rs!INTERNAL_EXTENSION)
                            .LocalNumber = rs!LOCAL_NUMBER
                        End With
                        rs.MoveNext
                        oMember.Phone.GotoNext
                    Loop
                Else
                    Do While Not rs.EOF
                        oMember.Phone.GotoFirst
                        bFound = False
                        Do While Not oMember.Phone.EOF
                            If Val(oMember.Phone.Type) = rs!PHONE_TYPE_CID Then
                                With oMember.Phone
                                    .AreaCode = IIf(IsNull(rs!AREA_CODE), "", rs!AREA_CODE)
                                    .CountryCode = IIf(IsNull(rs!COUNTRY_CODE), "", rs!COUNTRY_CODE)
                                    .Extension = IIf(IsNull(rs!INTERNAL_EXTENSION), "", rs!INTERNAL_EXTENSION)
                                    .LocalNumber = IIf(IsNull(rs!LOCAL_NUMBER), "", rs!LOCAL_NUMBER)
                                End With
                                bFound = True
                                Exit Do
                            End If
                            oMember.Phone.GotoNext
                        Loop
                        If Not bFound Then
                            oMember.Phone.Type = rs!PHONE_TYPE_CID
                            With oMember.Phone
                                .AreaCode = IIf(IsNull(rs!AREA_CODE), "", rs!AREA_CODE)
                                .CountryCode = IIf(IsNull(rs!COUNTRY_CODE), "", rs!COUNTRY_CODE)
                                .Extension = IIf(IsNull(rs!INTERNAL_EXTENSION), "", rs!INTERNAL_EXTENSION)
                                .LocalNumber = IIf(IsNull(rs!LOCAL_NUMBER), "", rs!LOCAL_NUMBER)
                            End With
                        End If
                        rs.MoveNext
                    Loop
                End If
                    
            Case "MMI_SITE"
                If rs Is Nothing Then GoTo NextRec
                
                rs.MoveFirst
                
                If oMember.Facility Is Nothing Then
                    Do While Not rs.EOF
                        'Eic data is broken
                        If IsNull(rs!FACILITY_CID) Then GoTo NextRec
                            
                        oMember.Facility = rs!FACILITY_CID
                        With oMember.Facility
                            .MedRecNum = rs!MEDICAL_RECORD_NUMBER
                            '.Name = rs!Name
                            .RadiologyNum = rs!RADIOLOGY_NUMBER
                        End With
                        rs.MoveNext
                        oMember.Facility.GotoNext
                    Loop
                Else
                    
                    If IsNull(rs!FACILITY_CID) Then GoTo NextRec
                
                    Do While Not rs.EOF
                        oMember.Facility.GotoFirst
                        bFound = False
                        Do While Not oMember.Facility.EOF
                            If oMember.Facility.Type = rs!FACILITY_CID Then
                                With oMember.Facility
                                    .MedRecNum = rs!MEDICAL_RECORD_NUMBER
                                    '.Name = rs!Name
                                    .RadiologyNum = rs!RADIOLOGY_NUMBER
                                End With
                                bFound = True
                                Exit Do
                            End If
                            oMember.Facility.GotoNext
                        Loop
                        If Not bFound Then
                            oMember.Facility.Type = rs!FACILITY_CID
                            With oMember.Facility
                                .MedRecNum = rs!MEDICAL_RECORD_NUMBER
                                '.Name = rs!Name
                                .RadiologyNum = rs!RADIOLOGY_NUMBER
                            End With
                                
                        End If
                        rs.MoveNext
                    Loop
                End If
            
        End Select
NextRec:
    Next

    Set TransferFromRSobj = oDoc
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "CHCSII_MMI.TransferData.TransferFromRSobj", Err.Description, Err.HelpFile, Err.HelpContext
End Function
