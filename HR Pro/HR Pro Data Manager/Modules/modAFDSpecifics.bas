Attribute VB_Name = "modAFDSpecifics"
Option Explicit

' What type of AFD is installed
Public Enum AFDTypes
  AFD_Disabled = 0
  AFD_PostCode = 1
  AFD_PostCodeplus = 2
  AFD_NamesNumbers = 3
End Enum

' Global Variable (to hold if AFD Module is enabled)
Public gfAFDEnabled As AFDTypes

' AFD Names And Numbers Functions
Declare Function GetPostcode Lib "afdnn32.dll" (nnRec As nnRecType, ByVal RecNo As Long, ByVal nnFlags As Long) As Long

' AFD Postcode Plus Functions
Declare Function getPCPPostLocation Lib "pcpv232.dll" Alias "GetPostcode" (PcLoc As PcLocRec, ByVal RecNo As Long) As Long
Declare Function getPCPPostAddress Lib "pcpv232.dll" Alias "GetAddress" (PcAddr As PcAddrRec, ByVal RecNo As Long) As Long

' AFD Poscode Normal Functions
Declare Function getPostcodeFirst Lib "PCODE32.DLL" Alias "GetFirst" (details As PostcodeData, Flags As Integer) As Long
Declare Function getPostcodeNext Lib "PCODE32.DLL" Alias "GetNext" (details As PostcodeData) As Long


''' STUFF COPIED FROM EXAMPLES (probably most of it is not used by us)

'Declare the Names & Numbers Address Record
Type nnRecType
 RevisionID As String * 2       'Revision number of Names & Numbers - set to '02'
 PostCode As String * 8         'Postcode
 PostcodeFrom As String * 8     'Postcode From (used where a range of postcodes is required & also returns the OLD postcode if a changed postcode is detected)
 DPS As String * 2              'Delivery Point Suffix
 MailSort As String * 5         'Mailsort Code
 StdCode As String * 8          'Predicted STD Code for Postal Sector
 WardCode As String * 6         'Local Authority Ward Code
 WardName As String * 30        'Local Authority Ward Name
 NHSAreaCode As String * 3      'NHS Area - Code
 NHSAreaName As String * 50     'NHS Area - Name
 NHSRegionCode As String * 3    'NHS Region - Code
 NHSRegionName As String * 40   'NHS Region - Name
 PostcodeType As String * 6     'L=Large User, S=Small User, N=Non PAF
 GridE As String * 6            'Grid Reference Easting
 GridN As String * 6            'Grid Reference Northing
 Distance As String * 6         'Distance from Test Grid Ref - in 10ths of a Kilometer
 PostalCounty As String * 20          'County Name according to Postal Authority
 AdministrativeCounty As String * 20  'County Name for Local Authority purposes
 TraditionalCounty As String * 20     'Traditional County Name
 Town As String * 30            'Post Town
 Locality As String * 70        'Locality (includes Double Dependant Locality)
 Street As String * 120         'Street or Thoroughfare (includes Dependant Thoroughfare)
 HouseNo As String * 10         'Building Number
 Building As String * 60        'Building Name
 SubBuilding As String * 60     'Sub-Building Name
 Phone As String * 20           'Telephone No where known, incl STD Code
 Surname As String * 30         'Surname
 FirstName As String * 30       'First Name
 Initial2 As String * 6         'Initial of Second Forename
 Residency As String * 6        'Length of time resident at this address
 Gender As String * 6           'M=Male;  F=Female; X=Ambiguous; Blank=Not Known
 Organisation As String * 120    'Organisation Name (includes Department, if any)
 Business As String * 100        'Business Description - eg 'Solicitors'
 Size As String * 6             'No. of employees of the business '*
 SIC As String * 10             'Standard Industry Classification
 TVRegion As String * 30        'TV Region '*
 Constituency As String * 50    'Parliamentary Constituency '*
 Authority As String * 50       ' Local Authority / Unitary Authority '*
 CameoUKCategory As String * 2              'Cameo Geodemographic data - UK Category
 CameoIncomeCategory As String * 2          'Income
 CameoInvestorCategory As String * 2        'Investor
 CameoFinancialCategory As String * 2       'Financial
 CameoUnEmploymentCategory As String * 2    'Unemployment Category
 CameoPropertyCategory As String * 2        'Property
 reserved As String * 231       'Reserved for future use
End Type

' Define the AFD PostCode Normal data structure
Type PostcodeData
  lookup As String * 60
  PostCode As String * 8
  PostcodeType As String * 2
  Organisation As String * 30
  property As String * 30
  Street As String * 60
  Locality As String * 60
  Town As String * 30
  County As String * 30
  CountyOption As String * 1
  MailSort As String * 5
  StdCode As String * 8
  GridN As String * 6
  GridE As String * 5
  reserved As String * 50
End Type

' Define the AFD PostCode Plus Location data structure
Type PcLocRec
     Vli As String * 4             'Variable Length Indicator- Indicates length of record to DLL function - for future development
     PostCode  As String * 8
     DepStreet  As String * 60     'Equivalent to Royal Mail 'Dependent Thoroughfare' + Descriptor field
     Street  As String * 60        'Equivalent to Royal Mail 'Thoroughfare' + Descriptor field
     DblDepLoc  As String * 35     'Equivalent to Royal Mail 'Double Dependent Locality' field
     DepLoc  As String * 35        'Equivalent to Royal Mail 'Dependent Locality' field
     Town  As String * 30
     CountySource As String * 1    ' '0' if County Data Supplied by Royal Mail, '1' if supplied by AFD
     County  As String * 30
     CountyAbbr  As String * 20    'Abbreviated county - where available
     MailSort  As String * 5       'Used to sort post for the Mailsort Discount scheme
     StdCode  As String * 8        'Our best guess at the STD code for the postcode
     GridE  As String * 5          'Grid Reference Easting
     GridN  As String * 5          'Grid Reference Northing
     WardCode  As String * 6       'Local Authority Ward Code
     WardName  As String * 50      'Local Authority Ward Name
     NHSCode  As String * 3        'Health Service Code
     NHSName  As String * 50       'Health Service Name
     NHSRegionCode  As String * 3  'Health Service Region Code
     NHSRegion  As String * 40     'Health Service Region Name
     DeliveryPoints  As String * 3 'Number of Addresses on a postcode locality
     UserType  As String * 1       'L=Large User   S=Small User
     Constituency As String * 50   'Parliamnetary Constituency
     TVRegion As String * 30       'Television Region (Not TV Station)
     Authority As String * 50      'Local / Unitary Authority
     TraditionalCounty As String * 30    ' Traditional County
     AdministrativeCounty As String * 30 ' Administrative County
End Type

' Define the AFD PostCode Plus Address data structure
Type PcAddrRec
     Vli As String * 4              'Indicates length of record to DLL function - for future development
     POBox As String * 6            'PO Box Number - means there will be no Builing or Street
     Organisation As String * 60    'Name of Organisation
     Department As String * 60      'Department of Organisation
     Number As String * 4           'Number of Building
     SubBuilding As String * 60     'Sub Building Name/Number eg 'Flat 1' or '6A'
     Building As String * 60        'Building Name
     DPS As String * 2              'Delivery Point Suffix or 'Mailcode'
     Households As String * 4       'Number of households in a multi-occupancy address
     reserved As String * 64
End Type


'Declare the changes data type
'Type ChangesType
'     PostCode  As String * 8
'End Type

'Type PhoneChangesType
'  PhoneNo As String * 15
'End Type
'Global Declaration
'Global OldNumber As PhoneChangesType
'Global NewNumber As PhoneChangesType

'Declare the General String data type
'Type AFDStringType
'  Value As String * 255
'End Type

' AFD Names And Number General Working Data Records
'Global nnRec As nnRecType
'Global OldPC As ChangesType
'Global NewPC As ChangesType
'Global ListRec As AFDStringType
'Global FieldName As AFDStringType

'Revision number for AFD Names and Numbers
Global Const nnRevisionID$ = "02"

'Search Control Flag
'Bits 0-3 specify skipping
' Global Const NO_SKIP = 0
' Global Const SKIP_TO_NEXT_SECTOR = 1
' Global Const SKIP_TO_NEXT_OUTCODE = 2
' Global Const SKIP_TO_NEXT_POSTTOWN = 3
' Global Const SKIP_TO_NEXT_POSTCODE_AREA = 4

'Bits 4-6 specify data subsets to use
'(currently only USE_ALL_SUBSETS is available)
' Global Const USE_ALL_SUBSETS = 0

'Bit 8 specifies whether GetPostcode should return all records or
' only those which match the last search
 Global Const RETURN_ALL_RECORDS = 0
 Global Const RETURN_MATCHING_RECORDS = 256

'Bit 9 specifies whether to provide Sector breaks or not
' Global Const NO_SECTOR_BREAKS = 0
' Global Const RETURN_SECTOR_BREAKS = 512

'Return Codes
 Global Const INVALID_POSTCODE = -1
 Global Const POSTCODE_NOT_FOUND = -2
 Global Const INVALID_RECORD_NUMBER = -3
 Global Const ERROR_OPENING_FILES = -4
 Global Const FILE_READ_ERROR = -5
 Global Const END_OF_SEARCH = -6
 Global Const DATA_LICENSE_ERROR = -7
 Global Const CONFLICTING_SEARCH_PARAMETERS = -8  'ie trying to Search Org & Surname/Firstname at the same time


Sub ClearnnRec(nn As nnRecType)

'Initialises Names & Numbers record with space
' - and RevisionId with constant value

    nn.RevisionID = nnRevisionID$
    nn.PostCode = ""
    nn.PostcodeFrom = ""
    nn.DPS = ""
    nn.MailSort = ""
    nn.StdCode = ""
    nn.WardCode = ""
    nn.WardName = ""
    nn.NHSAreaCode = ""
    nn.NHSAreaName = ""
    nn.NHSRegionCode = ""
    nn.NHSRegionName = ""
    nn.PostcodeType = ""
    nn.GridN = ""
    nn.GridE = ""
    nn.Distance = ""
    nn.PostalCounty = ""
    nn.AdministrativeCounty = ""
    nn.TraditionalCounty = ""
    nn.Town = ""
    nn.Locality = ""
    nn.Street = ""
    nn.HouseNo = ""
    nn.Building = ""
    nn.SubBuilding = ""
    nn.Phone = ""
    nn.Surname = ""
    nn.FirstName = ""
    nn.Initial2 = ""
    nn.Residency = ""
    nn.Gender = ""
    nn.Organisation = ""
    nn.Business = ""
    nn.Size = ""
    nn.SIC = ""
    nn.TVRegion = ""
    nn.Constituency = ""
    nn.Authority = ""
    nn.CameoUKCategory = ""
    nn.CameoIncomeCategory = ""
    nn.CameoInvestorCategory = ""
    nn.CameoFinancialCategory = ""
    nn.CameoUnEmploymentCategory = ""
    nn.CameoPropertyCategory = ""
    nn.reserved = ""
End Sub

Sub ShowError(ErrorNumber&)
 Select Case ErrorNumber&
    Case -1: COAMsgBox "Invalid Postcode", 16
    Case -2: COAMsgBox "Postcode Not Found", 16
    Case -3: COAMsgBox "Invalid Record Number", 16
    Case -4: COAMsgBox "Error Opening Postcode Files", 16
    Case -5: COAMsgBox "File Read Error", 16
    Case -6: COAMsgBox "End of Search", 16
    Case -7: COAMsgBox "Data License Error", 16
    Case -8: COAMsgBox "Conflicting Search Parameters", 16      'Occurs when Organisation &
                                                                                            'Name Searches are attempted at same time
 End Select

Screen.MousePointer = vbDefault

End Sub

Public Sub modAfdShowMappedFields(TableID As Long, FieldName As String, PostCode As String, frmForm As Form)

On Error GoTo AdfShowMappedFieldsError

  Dim rs As ADODB.Recordset           'Recordset containing mapped fields (columnIDs)
  Dim sSQL As String            'source of recordset
  Dim fIndividual As Boolean    'individual or merged address fields

  'Let the user know something is happening
  Screen.MousePointer = vbHourglass
  
  'Load the Afd form
  Load frmAFDFields
  
  'Set the source of the recordset
  sSQL = "SELECT * from asrsyscolumns WHERE tableid = " & frmForm.TableID & _
  " AND columnname = '" & FieldName & "'"

  'Load recordset
  Set rs = datGeneral.GetRecords(sSQL)

  'Go through each field (different ones depending on the value of fIndividual)
  'If there is a valid column mapped then set the tag property of the relevant text
  'box on the Afd form, otherwise, disable the checkbox and the text field on the Afd form.
  If Not rs.BOF And Not rs.EOF Then
    fIndividual = rs.Fields("Afdindividual")
    If Not fIndividual Then
      
      If rs.Fields("Afdforename") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtMergedForename.Tag = rs.Fields("Afdforename")
      Else
        frmAFDFields.txtMergedForename.Tag = 0
        frmAFDFields.chkMergedForename.Value = False
        frmAFDFields.chkMergedForename.Enabled = False
        frmAFDFields.txtMergedForename.Enabled = False
        frmAFDFields.txtMergedForename.BackColor = &H8000000F
      End If
            
      If rs.Fields("Afdinitial") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtMergedInitials.Tag = rs.Fields("Afdinitial")
      Else
        frmAFDFields.txtMergedInitials.Tag = 0
        frmAFDFields.chkMergedInitials.Value = False
        frmAFDFields.chkMergedInitials.Enabled = False
        frmAFDFields.txtMergedInitials.Enabled = False
        frmAFDFields.txtMergedInitials.BackColor = &H8000000F
      End If
            
      If rs.Fields("Afdsurname") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtMergedSurname.Tag = rs.Fields("Afdsurname")
      Else
        frmAFDFields.txtMergedSurname.Tag = 0
        frmAFDFields.chkMergedSurname.Value = False
        frmAFDFields.chkMergedSurname.Enabled = False
        frmAFDFields.txtMergedSurname.Enabled = False
        frmAFDFields.txtMergedSurname.BackColor = &H8000000F
      End If
      
      If rs.Fields("Afdaddress") <> 0 Then
        frmAFDFields.txtMergedAddress.Tag = rs.Fields("Afdaddress")
      Else
        frmAFDFields.txtMergedAddress.Tag = 0
        frmAFDFields.chkMergedAddress.Value = False
        frmAFDFields.chkMergedAddress.Enabled = False
        frmAFDFields.txtMergedAddress.Enabled = False
        frmAFDFields.txtMergedAddress.BackColor = &H8000000F
      End If
      
      If rs.Fields("Afdtelephone") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
      frmAFDFields.txtMergedTelephone.Tag = rs.Fields("Afdtelephone")
      Else
        frmAFDFields.txtMergedTelephone.Tag = 0
        frmAFDFields.chkMergedTelephone.Value = False
        frmAFDFields.chkMergedTelephone.Enabled = False
        frmAFDFields.txtMergedTelephone.Enabled = False
        frmAFDFields.txtMergedTelephone.BackColor = &H8000000F
      End If
      
    Else
      
      If rs.Fields("Afdforename") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtForename.Tag = rs.Fields("Afdforename")
      Else
        frmAFDFields.txtForename.Tag = 0
        frmAFDFields.chkForename.Value = False
        frmAFDFields.chkForename.Enabled = False
        frmAFDFields.txtForename.Enabled = False
        frmAFDFields.txtForename.BackColor = &H8000000F
      End If
            
      If rs.Fields("Afdinitial") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtInitials.Tag = rs.Fields("Afdinitial")
      Else
        frmAFDFields.txtInitials.Tag = 0
        frmAFDFields.chkInitials.Value = False
        frmAFDFields.chkInitials.Enabled = False
        frmAFDFields.txtInitials.Enabled = False
        frmAFDFields.txtInitials.BackColor = &H8000000F

      End If
            
      If rs.Fields("Afdsurname") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtSurname.Tag = rs.Fields("Afdsurname")
      Else
        frmAFDFields.txtSurname.Tag = 0
        frmAFDFields.chkSurname.Value = False
        frmAFDFields.chkSurname.Enabled = False
        frmAFDFields.txtSurname.Enabled = False
        frmAFDFields.txtSurname.BackColor = &H8000000F

      End If
      
      If rs.Fields("Afdproperty") <> 0 Then
        frmAFDFields.txtProperty.Tag = rs.Fields("Afdproperty")
      Else
        frmAFDFields.txtProperty.Tag = 0
        frmAFDFields.chkProperty.Value = False
        frmAFDFields.chkProperty.Enabled = False
        frmAFDFields.txtProperty.Enabled = False
        frmAFDFields.txtProperty.BackColor = &H8000000F

      End If
      
      If rs.Fields("Afdstreet") <> 0 Then
        frmAFDFields.txtStreet.Tag = rs.Fields("Afdstreet")
      Else
        frmAFDFields.txtStreet.Tag = 0
        frmAFDFields.chkStreet.Value = False
        frmAFDFields.chkStreet.Enabled = False
        frmAFDFields.txtStreet.Enabled = False
        frmAFDFields.txtStreet.BackColor = &H8000000F

      End If
      
      If rs.Fields("Afdlocality") <> 0 Then
        frmAFDFields.txtLocality.Tag = rs.Fields("Afdlocality")
      Else
        frmAFDFields.txtLocality.Tag = 0
        frmAFDFields.chkLocality.Value = False
        frmAFDFields.chkLocality.Enabled = False
        frmAFDFields.txtLocality.Enabled = False
        frmAFDFields.txtLocality.BackColor = &H8000000F

      End If
      
      If rs.Fields("Afdtown") <> 0 Then
        frmAFDFields.txtTown.Tag = rs.Fields("Afdtown")
      Else
        frmAFDFields.txtTown.Tag = 0
        frmAFDFields.chkTown.Value = False
        frmAFDFields.chkTown.Enabled = False
        frmAFDFields.txtTown.Enabled = False
        frmAFDFields.txtTown.BackColor = &H8000000F

      End If
      
      If rs.Fields("Afdcounty") <> 0 Then
        frmAFDFields.txtCounty.Tag = rs.Fields("Afdcounty")
      Else
        frmAFDFields.txtCounty.Tag = 0
        frmAFDFields.chkCounty.Value = False
        frmAFDFields.chkCounty.Enabled = False
        frmAFDFields.txtCounty.Enabled = False
        frmAFDFields.txtCounty.BackColor = &H8000000F

      End If
      
      If rs.Fields("Afdtelephone") <> 0 And gfAFDEnabled = AFD_NamesNumbers Then
        frmAFDFields.txtTelephone.Tag = rs.Fields("Afdtelephone")
      Else
        frmAFDFields.txtTelephone.Tag = 0
        frmAFDFields.chkTelephone.Value = False
        frmAFDFields.chkTelephone.Enabled = False
        frmAFDFields.txtTelephone.Enabled = False
        frmAFDFields.txtTelephone.BackColor = &H8000000F

      End If
      
    End If
    'Clear recordset reference
    Set rs = Nothing
  Else
    'Here is no data is found in the recordset...should never happen, but just incase
    Set rs = Nothing
    Exit Sub
  End If

  'Call the Afd routines. If they fail, exit sub, if not, show the Afd form
  If frmAFDFields.InitialiseAFD(PostCode, fIndividual, frmForm, FieldName) = False Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  'Return mousepointer to normal
  Screen.MousePointer = vbDefault
  
  'Show the Afd form
  frmAFDFields.Show vbModal

AdfShowMappedFieldsResume:

Exit Sub

AdfShowMappedFieldsError:

COAMsgBox "Error : " & Err.Number & " - " & Err.Description & " - ModAFDSpecifics.modAfdShowMappedFields", vbOKOnly, "Error"
Resume AdfShowMappedFieldsResume

End Sub

Public Function GetAFDPostcode(ByRef pstrPostcode As String, ByVal plngStartRecno As Long, ByVal plngReturnSet As Long) As Long

'Wrapper function for the AFD Postcode types

Dim lngResult As Long
Dim iCountLocations As Integer
Dim lngAddressResult As Long
Dim lngLocationResult As Long
Dim iCount As Integer
Dim bRecordFound As Boolean

' Variables to hold the postcode data to transfer to AFD dlls
Dim oPostCodeNormal As PostcodeData
Dim oNamesAndNumbers As nnRecType
Dim oPostCodePlusLocation(12) As PcLocRec
Dim oPostCodePlusAddress As PcAddrRec

lngResult = 0

' AFD Names and Number routine
If gfAFDEnabled = AFD_NamesNumbers Then
  ClearnnRec oNamesAndNumbers
  oNamesAndNumbers.PostCode = pstrPostcode
  lngResult = GetPostcode(oNamesAndNumbers, plngStartRecno, plngReturnSet)
  If Not lngResult < 0 Then
    With oPostCode
      .PostCode = oNamesAndNumbers.PostCode
      .FirstName = oNamesAndNumbers.FirstName
      .Initial2 = oNamesAndNumbers.Initial2
      .Surname = oNamesAndNumbers.Surname
      .Building = oNamesAndNumbers.Building
      .HouseNo = oNamesAndNumbers.HouseNo
      .Street = oNamesAndNumbers.Street
      .Locality = oNamesAndNumbers.Locality
      .Town = oNamesAndNumbers.Town
      .County = oNamesAndNumbers.PostalCounty
      .Phone = oNamesAndNumbers.Phone
      .Organisation = oNamesAndNumbers.Organisation
    End With
  End If
End If

' AFD Postcode Normal routine
If gfAFDEnabled = AFD_PostCode Then
  oPostCodeNormal.lookup = pstrPostcode
  oPostCodeNormal.CountyOption = "6"  ' Set "Traditional" county results
  lngResult = getPostcodeFirst(oPostCodeNormal, 0)
  
  ' Cannot go straight to desired record, have to loop through
  For iCount = 1 To plngStartRecno - 1
    lngResult = getPostcodeNext(oPostCodeNormal)
  Next iCount
  
  If Not lngResult < 0 Then
    With oPostCode
      .PostCode = oPostCodeNormal.PostCode
      .FirstName = ""
      .Initial2 = ""
      .Surname = ""
      .Building = ""
      .HouseNo = oPostCodeNormal.property
      .Street = oPostCodeNormal.Street
      .Locality = oPostCodeNormal.Locality
      .Town = oPostCodeNormal.Town
      .County = oPostCodeNormal.County
      .Phone = ""
      .Organisation = ""
      lngResult = 10       'Fool HR-Pro into thinking there is more than one postcode
    End With
  End If
End If

' AFD Postcode Plus routine
If gfAFDEnabled = AFD_PostCodeplus Then
  
  ' Calculate the total amount of locations
  oPostCodePlusLocation(1).PostCode = pstrPostcode
  oPostCodePlusLocation(1).Vli = Format$(Len(oPostCodePlusLocation(1)), "0000")
  lngLocationResult = getPCPPostLocation(oPostCodePlusLocation(1), 1)
  bRecordFound = False
  
  ' Populate each of the location records
  For iCountLocations = 1 To lngLocationResult
    oPostCodePlusLocation(iCountLocations).PostCode = pstrPostcode
    oPostCodePlusLocation(iCountLocations).Vli = Format$(Len(oPostCodePlusLocation(iCountLocations)), "0000")
    lngLocationResult = getPCPPostLocation(oPostCodePlusLocation(iCountLocations), iCountLocations)
  
    'Calculate total amount of addresses found.
    lngResult = lngResult + oPostCodePlusLocation(iCountLocations).DeliveryPoints
 
    ' Get first address of the current location (if we are looking for an address outside of this location, then jump...
    If Not bRecordFound Then
     If plngStartRecno <= Val(oPostCodePlusLocation(iCountLocations).DeliveryPoints) Then
       oPostCodePlusAddress.Vli = Format$(Len(oPostCodePlusAddress), "0000")
       lngAddressResult = getPCPPostAddress(oPostCodePlusAddress, plngStartRecno)
  
       If Not lngAddressResult < 0 Then
         With oPostCode
           .PostCode = pstrPostcode
           .FirstName = ""
           .Initial2 = ""
           .Surname = ""
           .Building = Trim(Trim(oPostCodePlusAddress.Building) & " " & oPostCodePlusAddress.SubBuilding)
           .HouseNo = oPostCodePlusAddress.Number
           .Street = oPostCodePlusLocation(iCountLocations).Street
           .Locality = oPostCodePlusLocation(iCountLocations).DepLoc
           .Town = oPostCodePlusLocation(iCountLocations).Town
           .County = oPostCodePlusLocation(iCountLocations).TraditionalCounty
           .Phone = ""
           .Organisation = ""
           bRecordFound = True
         End With
       End If
     Else
       ' Outside of this location, calculate the number in the next location
       plngStartRecno = plngStartRecno - oPostCodePlusLocation(iCountLocations).DeliveryPoints
     End If
    End If

  Next iCountLocations
End If

' Return the result
GetAFDPostcode = lngResult

End Function

' What version of AFD is installed on this machine
Public Function AFDVersion() As AFDTypes

  Dim lngResult As Long
  Dim oPostCodeNormal As PostcodeData
  Dim oNamesAndNumbers As nnRecType
  Dim oPostCodePlus(12) As PcLocRec
  Dim lngAFDType As AFDTypes
  Dim strSeedAFDNormal As String
  Dim strSeedAFDPlus As String
  Dim strSeedAFDNN As String

  If GetSystemSetting("Development", "AFD_Evaluation_Enable", False) Then
    strSeedAFDNormal = GetSystemSetting("Development", "AFD_Evaluation_Seed_Normal", "B13 9JD")
    strSeedAFDPlus = GetSystemSetting("Development", "AFD_Evaluation_Seed_Plus", "B45 9AA")
    strSeedAFDNN = GetSystemSetting("Development", "AFD_Evaluation_Seed_NN", "IS2 9NY")
  Else
    strSeedAFDNormal = "HP2 6LQ"
    strSeedAFDPlus = "HP2 6LQ"
    strSeedAFDNN = "HP2 6LQ"
  End If
  
  lngResult = -1
  lngAFDType = AFD_Disabled

  ' Trap if AFD is not installed
  On Error GoTo AfdNotInstalled:
    
  ' Check for AFD Names and numbers
  ClearnnRec oNamesAndNumbers
  'oNamesAndNumbers.PostCode = IIf(ASRDEVELOPMENT, "IS2 9NY", "HP2 6LQ")
  oNamesAndNumbers.PostCode = strSeedAFDNN
  lngResult = GetPostcode(oNamesAndNumbers, 1, RETURN_ALL_RECORDS)
  
  If Not lngResult < 0 Then
    lngAFDType = AFD_NamesNumbers
  Else
    ' Check for AFD Postcode Plus
    'oPostCodePlus(1).PostCode = IIf(ASRDEVELOPMENT, "B45 9AA", "HP2 6LQ")
    oPostCodePlus(1).PostCode = strSeedAFDPlus
    oPostCodePlus(1).Vli = Format$(Len(oPostCodePlus(1)), "0000")
    lngResult = getPCPPostLocation(oPostCodePlus(1), 1)
    
    If Not lngResult < 0 Then
      lngAFDType = AFD_PostCodeplus
    Else
      'Check AFD Postcode Normal
      'oPostCodeNormal.lookup = IIf(ASRDEVELOPMENT, "B13 9JD", "HP2 6LQ")
      oPostCodeNormal.lookup = strSeedAFDNormal
      lngResult = getPostcodeFirst(oPostCodeNormal, 0)
      
      If Not lngResult < 0 Then
        lngAFDType = AFD_PostCode
      End If
    End If
  End If
  
  ' Set the return type
  AFDVersion = lngAFDType
  
  Exit Function

AfdNotInstalled:
  lngResult = -1
  Resume Next

End Function
