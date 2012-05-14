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
Attribute VB_Name = "Immunization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarRouteNCID As String 'local copy
Private mvarsRouteName As String 'local copy
Private mvarSiteNCID As String 'local copy
Private mvarfDose As Single 'local copy
Private mvarResult As String 'local copy
Private mvarNextDueDate As Date 'local copy
Private mvarsStatusName As String 'local copy
Private mvarsStatusNCID As String 'local copy
Private mvardAdminDate As Date 'local copy
Private mvarsEventID As String 'local copy
Private mvarsManufacturerNCID As String 'local copy
Private mvarsManufacturerName As String 'local copy
Private mvarsLot As String 'local copy
Private mvardExpirationDate As Date 'local copy
Private mvarsAdminLocation As String 'local copy
Private mvarsAdminLocationNCID As String 'local copy
Private mvarsSeqNumberName As String 'local copy
Private mvarsVaccineNCID As String 'local copy
Private mvarsVaccineName As String 'local copy
Private mvarsUnitsNCID As String 'local copy
Private mvarsUnitsName As String 'local copy
Private mvarsClinicianSSN As String 'local copy
Private mvarsComment As String 'local copy
Private mvarsFMP As String 'local copy
Private mvarsSSN As String 'local copy
Private mvarIEN As Long 'local copy
Private mvarsClinicianName As String 'local copy
Private mvarsClinicianNCID As String 'local copy
Private mvarsEncounterNumber As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsSeqNumberNCID As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsSiteName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsCodedCommentsNCID As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsEditFlag As String 'local copy
'local variable(s) to hold property value(s)
Private mvarfReactionSize As Single 'local copy
'local variable(s) to hold property value(s)
Private mvardExemptionExpiration As Date 'local copy
Private mvarsExemptionStatus As String 'local copy
'local variable(s) to hold property value(s)
Private mvardTBPlantDate As Date 'local copy
Public Property Let dTBPlantDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTBPlantDate = 5
    mvardTBPlantDate = vData
End Property


Public Property Get dTBPlantDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTBPlantDate
    dTBPlantDate = mvardTBPlantDate
End Property



Public Property Let sExemptionStatus(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sExemptionStatus = 5
    mvarsExemptionStatus = vData
End Property


Public Property Get sExemptionStatus() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sExemptionStatus
    sExemptionStatus = mvarsExemptionStatus
End Property



Public Property Let dExemptionExpiration(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dExemptionExpiration = 5
    mvardExemptionExpiration = vData
End Property


Public Property Get dExemptionExpiration() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dExemptionExpiration
    dExemptionExpiration = mvardExemptionExpiration
End Property



Public Property Let fReactionSize(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fReactionSize = 5
    mvarfReactionSize = vData
End Property


Public Property Get fReactionSize() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fReactionSize
    fReactionSize = mvarfReactionSize
End Property



Public Property Let sEditFlag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEditFlag = 5
    mvarsEditFlag = vData
End Property


Public Property Get sEditFlag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEditFlag
    sEditFlag = mvarsEditFlag
End Property



Public Property Let sCodedCommentsNCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodedCommentsNCID = 5
    mvarsCodedCommentsNCID = vData
End Property


Public Property Get sCodedCommentsNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodedCommentsNCID
    sCodedCommentsNCID = mvarsCodedCommentsNCID
End Property



Public Property Let sManufacturerName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sManufacturerName = 5
    mvarsManufacturerName = vData
End Property


Public Property Get sManufacturerName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sManufacturerName
    sManufacturerName = mvarsManufacturerName
End Property



Public Property Let sSiteName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiteName = 5
    mvarsSiteName = vData
End Property


Public Property Get sSiteName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiteName
    sSiteName = mvarsSiteName
End Property



Public Property Let sSeqNumberNCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSeqNumberNCID = 5
    mvarsSeqNumberNCID = vData
End Property


Public Property Get sSeqNumberNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSeqNumberNCID
    sSeqNumberNCID = mvarsSeqNumberNCID
End Property



Public Property Let sStatusName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sStatusName = 5
    mvarsStatusName = vData
End Property


Public Property Get sStatusName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sStatusName
    sStatusName = mvarsStatusName
End Property



Public Property Let sUnitsName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnitsName = 5
    mvarsUnitsName = vData
End Property


Public Property Get sUnitsName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnitsName
    sUnitsName = mvarsUnitsName
End Property



Public Property Let sRouteName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sRouteName = 5
    mvarsRouteName = vData
End Property


Public Property Get sRouteName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sRouteName
    sRouteName = mvarsRouteName
End Property



Public Property Let sAdminLocationNCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAdminLocationNCID = 5
    mvarsAdminLocationNCID = vData
End Property


Public Property Get sAdminLocationNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAdminLocationNCID
    sAdminLocationNCID = mvarsAdminLocationNCID
End Property



Public Property Let sEncounterNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEncounterNumber = 5
    mvarsEncounterNumber = vData
End Property


Public Property Get sEncounterNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEncounterNumber
    sEncounterNumber = mvarsEncounterNumber
End Property



Public Property Let sClinicianNCID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClinicianNCID = 5
    mvarsClinicianNCID = vData
End Property


Public Property Get sClinicianNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClinicianNCID
    sClinicianNCID = mvarsClinicianNCID
End Property



Public Property Let sClinicianName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClinicianName = 5
    mvarsClinicianName = vData
End Property


Public Property Get sClinicianName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClinicianName
    sClinicianName = mvarsClinicianName
End Property



Public Property Let sVaccineName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVaccineName = 5
    mvarsVaccineName = vData
End Property


Public Property Get sVaccineName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVaccineName
    sVaccineName = mvarsVaccineName
End Property



Public Property Let lIEN(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IEN = 5
    mvarIEN = vData
End Property


Public Property Get lIEN() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IEN
    lIEN = mvarIEN
End Property



Public Property Let sSSN(ByVal vData As String)
Attribute sSSN.VB_Description = "SSN of patient (or sponsor)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSSN = 5
    mvarsSSN = vData
End Property


Public Property Get sSSN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSSN
    sSSN = mvarsSSN
End Property



Public Property Let sFMP(ByVal vData As String)
Attribute sFMP.VB_Description = "Family Member Prefix of person being immunized"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFMP = 5
    mvarsFMP = vData
End Property


Public Property Get sFMP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFMP
    sFMP = mvarsFMP
End Property



Public Property Let sComment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sComment = 5
    mvarsComment = vData
End Property


Public Property Get sComment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sComment
    sComment = mvarsComment
End Property



Public Property Let sClinicianSSN(ByVal vData As String)
Attribute sClinicianSSN.VB_Description = "SSN of clinician who gave immunization.\r\nIf empty, immunization is a transcribed (aka history) event."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClinicianSSN = 5
    mvarsClinicianSSN = vData
End Property


Public Property Get sClinicianSSN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClinicianSSN
    sClinicianSSN = mvarsClinicianSSN
End Property



Public Property Let sUnitsNCID(ByVal vData As String)
Attribute sUnitsNCID.VB_Description = "Units used to measure dosage of shot (ml, mg, cc, etc)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnitsNCID = 5
    mvarsUnitsNCID = vData
End Property


Public Property Get sUnitsNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnitsNCID
    sUnitsNCID = mvarsUnitsNCID
End Property



Public Property Let sVaccineNCID(ByVal vData As String)
Attribute sVaccineNCID.VB_Description = "NCID of Vaccine adminstered."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVaccineNCID = 5
    mvarsVaccineNCID = vData
End Property


Public Property Get sVaccineNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVaccineNCID
    sVaccineNCID = mvarsVaccineNCID
End Property



Public Property Let sSeqNumberName(ByVal vData As String)
Attribute sSeqNumberName.VB_Description = "Number in a series of shots for an immunization"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSeqNumber = 5
    mvarsSeqNumberName = vData
End Property


Public Property Get sSeqNumberName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSeqNumber
    sSeqNumberName = mvarsSeqNumberName
End Property



Public Property Let sAdminLocationName(ByVal vData As String)
Attribute sAdminLocationName.VB_Description = "Name of location where immunization administered."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAdminLocation = 5
    mvarsAdminLocation = vData
End Property


Public Property Get sAdminLocationName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAdminLocation
    sAdminLocationName = mvarsAdminLocation
End Property



Public Property Let dExpirationDate(ByVal vData As Date)
Attribute dExpirationDate.VB_Description = "Expiration date of vaccine"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dExpirationDate = 5
    mvardExpirationDate = vData
End Property


Public Property Get dExpirationDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dExpirationDate
    dExpirationDate = mvardExpirationDate
End Property



Public Property Let sLot(ByVal vData As String)
Attribute sLot.VB_Description = "Lot number of drug administered."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLot = 5
    mvarsLot = vData
End Property


Public Property Get sLot() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLot
    sLot = mvarsLot
End Property



Public Property Let sManufacturerNCID(ByVal vData As String)
Attribute sManufacturerNCID.VB_Description = "Manufacturer of drug being administered."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sManufacturer = 5
    mvarsManufacturerNCID = vData
End Property


Public Property Get sManufacturerNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sManufacturer
    sManufacturerNCID = mvarsManufacturerNCID
End Property



Public Property Let sEventID(ByVal vData As String)
Attribute sEventID.VB_Description = "CDR key for immunization"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEventID = 5
    mvarsEventID = vData
End Property


Public Property Get sEventID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEventID
    sEventID = mvarsEventID
End Property



Public Property Let dAdminDate(ByVal vData As Date)
Attribute dAdminDate.VB_Description = "Date vaccination was administered"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAdminDate = 5
    mvardAdminDate = vData
End Property


Public Property Get dAdminDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dAdminDate
    dAdminDate = mvardAdminDate
End Property



Public Property Let sStatusNCID(ByVal vData As String)
Attribute sStatusNCID.VB_Description = "Status of immunization (given, refused, invalid)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sStatusName = 5
    mvarsStatusNCID = vData
End Property


Public Property Get sStatusNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sStatusName
    sStatusNCID = mvarsStatusNCID
End Property



Public Property Let dNextDueDate(ByVal vData As Date)
Attribute dNextDueDate.VB_Description = "Due Date of next shot in series"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NextDueDate = 5
    mvarNextDueDate = vData
End Property


Public Property Get dNextDueDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NextDueDate
    dNextDueDate = mvarNextDueDate
End Property



Public Property Let sResult(ByVal vData As String)
Attribute sResult.VB_Description = "Result of TB Test.  ""Positive"", ""Negative"", or empty."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Result = 5
    mvarResult = vData
End Property


Public Property Get sResult() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Result
    sResult = mvarResult
End Property



Public Property Let fDose(ByVal vData As Single)
Attribute fDose.VB_Description = "size of dose of vaccine"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fDose = 5
    mvarfDose = vData
End Property


Public Property Get fDose() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fDose
    fDose = mvarfDose
End Property



Public Property Let sSiteNCID(ByVal vData As String)
Attribute sSiteNCID.VB_Description = "NCID for location on patient's body where shot administered."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SiteNCID = 5
    mvarSiteNCID = vData
End Property


Public Property Get sSiteNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SiteNCID
    sSiteNCID = mvarSiteNCID
End Property



Public Property Let sRouteNCID(ByVal vData As String)
Attribute sRouteNCID.VB_Description = "NCID for tracking route of drug (shot, pill, etc.)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RouteNCID = 5
    mvarRouteNCID = vData
End Property


Public Property Get sRouteNCID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RouteNCID
    sRouteNCID = mvarRouteNCID
End Property



