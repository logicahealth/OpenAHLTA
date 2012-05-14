VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CardReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const PDF417_Bfr_Size = 1024        ' PDF417 > 1020
Const PDF417_RCV_THRESH = 300

'       The thresholds are the number of characters between interrupts
'       to check to see if data is still coming in.  If it is short, more
'       [unnecessary] interrupts are generated.  If it is long, there
'       may be an [unnecessarily] long wait after the last character is
'       received.  At 9600 baud, there is about a millisecond between
'       characters, or a thousand bytes takes about a second.  The timer
'       function is a background service and not very precise.  Because
'       the clock resolution under Windows 3.1 is not very good (55 ms),
'       the threshold constants should not be set to less than 100 (ms).

'   *** Code 39 Bar Code Data Fields (Current Version)
  
Const c39CardType = 1

Const c39SSN = c39CardType + 1
Const c39DDS = c39SSN

Const c39SponSSN = c39SSN + 6
Const c39PayGrade = c39SponSSN
Const c39Status = c39SponSSN

Const c39SvcComp = c39SponSSN + 6
Const c39IssDate = c39SvcComp + 1
Const c39SecCode = c39IssDate + 3
 
'   *** PDF417 Bar Code Header Fields
Const PDF417_Hdr_Size = 9    ' Length of PDF417 Header
Const PDF417_ID_Code = 1     ' Ident Code ("IDUS")
Const PDF417_Version = 5     ' Bar Code Version
Const PDF417_Size = 6        ' Bar Code Size (File size - 5)
              ' Byte 6 through end of file
Const PDF417_CheckSum = 8
Const PDF417_Text_Size = 9   ' Byte 10 to start of JPEG
Const PDF417_CurVersion = 3  ' Current (latest) version of PDF417
'   *** PDF417 Bar Code Data Fields (Current Version)
Const pdfSponFlag = 1  ' Sponsor Flag
Const pdfName = pdfSponFlag + 1          ' Cardholder Name
Const pdfSSN = pdfName + 27
Const pdfFSN = pdfSSN + 9
Const pdfSSN2 = pdfFSN + 1           ' Sponsor or dependent
Const pdfDDS = pdfSSN2 + 9

Const pdfHgt = pdfDDS + 2
Const pdfWgt = pdfHgt + 2
Const pdfHair = pdfWgt + 3
Const pdfEye = pdfHair + 2
Const pdfDOB = pdfEye + 2

Const pdfBenefits = pdfDOB + 9
Const pdfChEffDate = pdfBenefits + 5
Const pdfChExpDate = pdfChEffDate + 9
  
Const pdfCardType = pdfChExpDate + 9
Const pdfIssDate = pdfCardType + 6   ' Version 2 to 3 (+2)
Const pdfExpDate = pdfIssDate + 9
Const pdfSecCode = pdfExpDate + 9

Const pdfSvcComp = pdfSecCode + 4
Const pdfStatus = pdfSvcComp + 1
Const pdfBranch = pdfStatus + 6
Const pdfRank = pdfBranch + 5
Const pdfGrade = pdfRank + 6        ' Version 1 to 2 (+4)

Const pdfGeneva = pdfGrade + 4      ' Sponsor
Const pdfBlood = pdfGeneva + 3

Const pdfSponName = pdfGrade + 4      ' Family member
Const pdfSponSSN = pdfSponName + 27
Const pdfRelat = pdfSponSSN + 9
         
Private msError As String
Private msErrorNum As String
Private msErrorObj As Error

Private mCData As CardData

Private mBranchCode As Dictionary

Public Property Get cData() As CardData
    Set cData = mCData
End Property

'we use this property so that the consumer can see what the error message is if they want to inspect it
Public Property Get errorMessage() As String
    errorMessage = msError
End Property

'we use this property so that the consumer can see what the error number is if they want to inspect it
Public Property Get errorNumber() As String
    errorNumber = msErrorNum
End Property

Public Function LoadCAC39(BC_Data As String) As Boolean
On Error GoTo ErrorHandler
  
    mCData.SSN = Format$(Decrypt(Mid$(BC_Data, c39SSN, 6), 6), "000000000")
    mCData.SponsorSSN = mCData.SSN '& "20" 'Not sure why we are hard coding this value, need to talk to Mark to find out why - 3/10/03 BM
    
    'This is according to the "specs":
    'mCData.PersonDesignatorIdentifier = Decrypt(Mid$(BC_Data, 9, 7), 7)
    'mCData.PersonelCategory = Mid$(BC_Data, 15, 1)
    'mCData.BranchCode = Mid$(BC_Data, 16, 1)
    'mCData.PersonDesignatorTypeCode = Mid$(BC_Data, 17, 1)
    
    'This works:
    mCData.PersonDesignatorIdentifier = Decrypt(Mid$(BC_Data, 9, 7), 7)
    mCData.PersonelCategory = Mid$(BC_Data, 16, 1)
    mCData.BranchCode = Mid$(BC_Data, 17, 1)
    mCData.PersonDesignatorTypeCode = Mid$(BC_Data, 8, 1)
    
    mCData.CardInstanceIdentifier = Mid$(BC_Data, 18, 1)
  
    LoadCAC39 = True
    Exit Function
ErrorHandler:
    LoadCAC39 = False
End Function

Public Function LoadCACPDF417(BC_Data As String) As Boolean
On Error GoTo ErrorHandler
  
    mCData.VersionCode = Mid$(BC_Data, 1, 1)
    mCData.SSN = Format$(Decrypt(Mid$(BC_Data, 2, 6), 6), "000000000")
    mCData.SponsorSSN = mCData.SSN
    mCData.PersonDesignatorTypeCode = Mid$(BC_Data, 8, 1)
    mCData.PersonDesignatorIdentifier = Decrypt(Mid$(BC_Data, 9, 7), 7)
    
    mCData.FirstName = UCase(Mid$(BC_Data, 16, 20))
    mCData.Surname = UCase(Mid$(BC_Data, 36, 26))
    mCData.DateOfBirth = DaysToDate(Decrypt(Mid$(BC_Data, 62, 4), 4))
    mCData.PersonelCategory = Mid$(BC_Data, 66, 1)
    mCData.BranchCode = Mid$(BC_Data, 67, 1)
    mCData.PersonnelEntitlementConditionType = Mid$(BC_Data, 68, 2)
    mCData.Rank = Mid$(BC_Data, 70, 6)
    
    mCData.PayPlanCode = Mid$(BC_Data, 76, 2)
    mCData.PayPlanGradeCode = Mid$(BC_Data, 78, 2)
    
    mCData.CardIssueDate = DaysToDate(Decrypt(Mid$(BC_Data, 80, 4), 4))
    mCData.CardExpirationDate = DaysToDate(Decrypt(Mid$(BC_Data, 84, 4), 4))
    
    mCData.CardInstanceIdentifier = Mid$(BC_Data, 88, 1)
  
    LoadCACPDF417 = True
    Exit Function
ErrorHandler:
    LoadCACPDF417 = False

End Function

'-------------------------------------------------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------------------------------------------------
Public Function LoadCode39(BC_Data As String, BC_Data_Length As Integer) As Boolean
    Dim CardType As Integer
    Dim RFU As Integer
    Dim iYear As Integer
    Dim iDays As Long
    Dim iDate As String
    Dim Asc_A As Long
    Dim Asc_1 As Long
On Error GoTo ErrorHandler
  
    Asc_A = Asc("A")
    Asc_1 = Asc("1")
    CardType = Asc(Mid$(BC_Data, c39CardType, 1))
    
    If CardType >= Asc_1 And CardType <= Asc("9") Then       ' 24 chars: 1-9
        CardType = CardType - Asc_1 + 1
        RFU = 3
    ElseIf CardType >= Asc_A And CardType <= Asc("J") Then ' 21 chars:  A-J
        CardType = CardType - Asc_A + 1
        RFU = 0                 ' Reserved for Future Use deleted
    Else
        msError = "Code 39 Card Type - Can not read card"
        LoadCode39 = False
        Exit Function
    End If

    'End of Bar Code (3 fields)
    mCData.BranchCode = Mid$(BC_Data, RFU + c39SvcComp, 1)
   
    'Beginning of Bar Code (2 fields)
    Select Case CardType          ' The rest depends on card type
    
        'NOTE:  We are coding for the CAC format currently!
        Case 1, 2, 3, 9          ' AD/Reserve/Retiree sponsor
            mCData.SSN = Format$(Decrypt(Mid$(BC_Data, c39SSN, 6), 6), "000000000")
            mCData.SponsorSSN = mCData.SSN '& "20"
            mCData.PersonDesignatorIdentifier = Decrypt(Mid$(BC_Data, 9, 7), 7)
            mCData.PersonelCategory = Mid$(BC_Data, 15, 1)
            mCData.BranchCode = Mid$(BC_Data, 16, 1)
            mCData.PersonDesignatorTypeCode = Mid$(BC_Data, 17, 1)
            mCData.CardInstanceIdentifier = Mid$(BC_Data, 18, 1)
            
            'mCData.Rank = Mid$(BC_Data, c39PayGrade, 4)
        Case 4, 5, 6, 10  ' FMP not available
            'mvarFMP = ""
        Case 7, 8                   ' Dependent w/o SSN
            mCData.SponsorSSN = Format$(Decrypt(Mid$(BC_Data, c39SponSSN, 6), 6), "000000000") & Mid$(BC_Data, c39DDS, 2)
        Case Else
            msError = "* Code 39 * Card Type Error" ' Unknown
            LoadCode39 = False
            Exit Function
    End Select
  
    LoadCode39 = True

    Exit Function
ErrorHandler:
    LoadCode39 = False
End Function


Private Function LoadPDF417(BC_Data As String, BC_Data_Length As Integer) As Boolean
    Dim kHdr As Integer   ' Length of Header
    Dim CardVers As Integer
    Dim CardType As Integer
    Dim SponFlag As Integer
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    Dim sText As String
    Dim iCnt As Integer
    Dim sDDS As String
On Error GoTo ErrorHandler

    kHdr = PDF417_Hdr_Size
  
    If Left$(BC_Data, (PDF417_Version - PDF417_ID_Code)) <> "IDUS" Then
        msError = " * PDF417 * Identification Code Error"
        LoadPDF417 = False
        Exit Function
    End If
  
    CardVers = Val(Mid$(BC_Data, PDF417_Version, PDF417_Size - PDF417_Version))
    
    If CardVers < 1 Or CardVers > PDF417_CurVersion Then
        msError = " * PDF417 * Card Version Error"
        LoadPDF417 = False
        Exit Function
    End If
    
    SponFlag = Asc(Mid$(BC_Data, kHdr + pdfSponFlag, 1))    ' Save for later
    
    If SponFlag = 49 Then
        mCData.SponsorSSN = Mid$(BC_Data, kHdr + pdfSSN, 9) & Mid$(BC_Data, kHdr + pdfDDS, 2)
    Else
        mCData.SponsorSSN = Mid$(BC_Data, kHdr + pdfSponSSN, 9) & Mid$(BC_Data, kHdr + pdfDDS, 2)
    End If
    
    mCData.FirstName = Trim(Mid$(BC_Data, kHdr + pdfName, 27))
    mCData.SSN = Mid$(BC_Data, kHdr + pdfSSN, 9)
    sText = Mid$(BC_Data, kHdr + pdfDOB, 9)
    sYear = Left$(sText, 4)
    sMonth = Mid$(sText, 5, 3)
    sDay = Right$(sText, 2)
  
    Select Case Mid$(BC_Data, kHdr + pdfCardType, 4)
        Case "2ACT"
            CardType = 1
            mCData.Status = "AD"
            mCData.Rank = Mid$(BC_Data, kHdr + pdfRank, 6)
        Case "2RET"
            CardType = 2
            mCData.Status = "RET"
            mCData.Rank = Mid$(BC_Data, kHdr + pdfRank, 6)
        Case "2RES"
            CardType = 3
            mCData.Status = "RV"
            mCData.Rank = Mid$(BC_Data, kHdr + pdfRank, 6)
        Case "2RRT"
            CardType = 9
            mCData.Status = "RET"
            mCData.Rank = Mid$(BC_Data, kHdr + pdfRank, 6)
        Case "1173"
            If SponFlag = 49 Then
              CardType = 4              ' Sponsor
              mCData.Status = "AD"
            Else
              CardType = 5              ' Dependent
              mCData.Status = "DEP"
            End If
        Case Else
            msError = "* PDF417 * Card Type Error"
            LoadPDF417 = False
            Exit Function
    End Select
    
    Select Case CardVers
        Case 1, 2
            kHdr = kHdr - 2               ' Back out change for rest of the string
        Case 3        ' Version 3 increased Card Type from 4 to 6
        
        Case Else
            msError = "* PDF417 * Card Version Error"
            LoadPDF417 = False
            Exit Function
    End Select
    
    mCData.ServiceCode = Mid$(BC_Data, kHdr + pdfSvcComp, 1)
    '   *** Finish Determining the Code 39 Card Type
    If CardType = 4 Then                  ' Sponsor
        Select Case Mid$(BC_Data, kHdr + pdfSvcComp, 1)
            Case "3" To "7"
                CardType = 10                 ' ResRet TA
                mCData.Status = "RET"
        End Select
    End If
    
    If CardType = 5 Then                  ' Dependent
        Select Case Mid$(BC_Data, kHdr + pdfSvcComp, 1)
            Case "P", "R", "T", "V", "X", "Z", "2"
                CardType = 6                ' Inactive Reserve Dependent
                mCData.Status = "DEP"
        End Select              '  (could be 8)
    End If
    
    'rtv = WriteJPEG(BC_Data, BC_Data_Length)
    LoadPDF417 = True

    Exit Function
ErrorHandler:
    LoadPDF417 = False
    msError = Err.Description
    msErrorNum = Err.Number
    Set msErrorObj = Error
End Function

Function WriteJPEG(ByVal InData As String, ByVal Data_Length As Integer)
  Dim JPEG_Start As Integer
  Dim JPEG_Size As Integer
  Dim JPEG_File As Integer
  Dim JPEG_Data As String
  Dim Asc_3 As Long
  Dim Destination As String
  On Error GoTo Err_WriteJPEG
  Asc_3 = Asc("3")
  Destination = "C:\CITA\PDF417.jpg"

'   *** Calculate length of JPEG data from PDF417 header and extract data

  JPEG_Start = PDF417_Hdr_Size + Asc(Mid$(InData, PDF417_Text_Size, 1)) + 1
  JPEG_Size = Asc(Mid$(InData, PDF417_Size, 1)) * &H100
  JPEG_Size = JPEG_Size + Asc(Mid$(InData, PDF417_Size + 1, 1)) + (PDF417_Size - 1)
  JPEG_Size = JPEG_Size - JPEG_Start + 1       ' JPEG size
  JPEG_Data = Mid$(InData, JPEG_Start, JPEG_Size)
'   *** Remove destination file
  JPEG_File = FreeFile
  Open Destination For Output As JPEG_File
  Close JPEG_File

'   *** Open destination file to write to

  JPEG_File = FreeFile
  Open Destination For Binary As JPEG_File

'   *** Write out JPEG data

  Put JPEG_File, , JPEG_Data
  Close JPEG_File
 ' Me!imgPDF417.Requery
  WriteJPEG = JPEG_Size
Bye_WriteJpeg:
  Exit Function

Err_WriteJPEG:
  WriteJPEG = -1 * Err
  Resume Bye_WriteJpeg

End Function

Public Function Decrypt(inputStr As String, InputLen As Integer) As Double
    Dim BinNr As Double
    Dim b32chr As Integer
    Dim iChr As Integer
    Dim Asc_A As Integer: Asc_A = Asc("A")
    Dim Asc_0 As Integer: Asc_0 = Asc("0")
    
    BinNr = 0
    For iChr = 1 To InputLen
        b32chr = Asc(Mid$(inputStr, iChr, 1))
        If b32chr >= Asc_A Then
            b32chr = b32chr - Asc_A + 10
        Else
            b32chr = b32chr - Asc_0
        End If
    
        BinNr = BinNr * 32 + b32chr
    
    Next
    
    Decrypt = BinNr
    
End Function

Public Function DaysToDate(days As Long) As Date
    Dim v As Long
    
    v = DateValue("1/1/1000")
    
    DaysToDate = CDate(days + v)

End Function

Private Sub Class_Initialize()
        
    Set mBranchCode = New Dictionary
    
    Set mCData = New CardData


    'Branch Code / Service
    mBranchCode.Add "A", "USA"
    mBranchCode.Add "C", "USCG"
    mBranchCode.Add "D", "DoD"
    mBranchCode.Add "F", "USAF"
    mBranchCode.Add "H", "USPHS"
    mBranchCode.Add "M", "USMC"
    mBranchCode.Add "N", "USN"
    mBranchCode.Add "O", "NOAA"
    mBranchCode.Add "1", "Foreign Army"
    mBranchCode.Add "2", "Foreign Navy"
    mBranchCode.Add "3", "Foreign Marine Corps"
    mBranchCode.Add "4", "Foreign Air Force"
    mBranchCode.Add "X", "Other/Not Applicable"

    
End Sub

Private Sub Class_Terminate()
    Set mCData = Nothing
    
End Sub


Function LoadBCData(BC_Data As String, BC_Data_Length As Integer) As Boolean
  Dim RetVal As Integer
  Dim CardType As Integer
On Error GoTo ErrorHandler
  
    If BC_Data_Length = 0 Then
        CardType = 0               ' Time out or Abort
    ElseIf BC_Data_Length = 9 Then
        CardType = 11
    ElseIf BC_Data_Length = 19 Then 'CAC Card
        LoadCAC39 BC_Data
    ElseIf BC_Data_Length = 21 Or BC_Data_Length = 24 Then
        CardType = LoadCode39(BC_Data, BC_Data_Length)
    ElseIf BC_Data_Length > 500 Then
        CardType = LoadPDF417(BC_Data, BC_Data_Length)
    Else
        CardType = 11
    End If

    LoadBCData = True
    Exit Function
ErrorHandler:
    LoadBCData = False
End Function
