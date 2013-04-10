Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
<System.Runtime.InteropServices.ProgId("StandardReports_NET.StandardReports")> Public Class StandardReports
	' JDM - THIS CLASS IS NO LONGER USED
	
	' Have not taken it out because it'll mess-up the class ID. Next big version change maybe a good time...
	
	
	Private mstrHexColour_FormBackground As String
	Private mstrHexColour_FormText As String
	Private mstrHexColour_AbsenceTypesBackground As String
	
  Dim mastrAbsenceTypes(,) As String ' Store the absence types (redefined later as ???,2 so as to auto clear it)
	'0 = Contains the Absence Type
	'1 = Contains the Absence String
	'2 = Is this absence in list by default
	
	Dim mlngFilterID As Integer
	Dim mlngPickListID As Integer
	Dim mdFromDate As Date
	Dim mdToDate As Date
	Dim mstrRecordSelectionType As String
	Dim mastrSelectedAbsenceTypes() As String
	
	Dim mlngPersonnelRecordID As Integer
	Dim mstrRealSource As String
	Dim mlngStandardReportType As StandardReportType
	
	' Matches the utility number in dat manager
	Public Enum StandardReportType
		utlUndefined = 0
		utlAbsenceBreakdown = 15
		utlBradfordFactor = 16
	End Enum
	
	Private mastrPersonnelFieldList() As String
	Private mbIndividualRecord As Boolean
	
	Public WriteOnly Property StandardReportType_Renamed() As Object
		Set(ByVal Value As Object)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pstrValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Value = "15" Then
				mlngStandardReportType = StandardReportType.utlAbsenceBreakdown
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pstrValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Value = "16" Then
				mlngStandardReportType = StandardReportType.utlBradfordFactor
			End If
			
		End Set
	End Property
	
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' Connection object passed in from the asp page
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				gADOCon = New ADODB.Connection
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
			Else
				gADOCon = Value
			End If
			
		End Set
	End Property
	
	Public WriteOnly Property FilterID() As Object
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pvFilterID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mlngFilterID = Value
			End If
		End Set
	End Property
	
	Public WriteOnly Property PickListID() As Object
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pvPickListID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mlngPickListID = Value
			End If
		End Set
	End Property
	
	Public WriteOnly Property FromDate() As Object
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsDate(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pvFromDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mdFromDate = Value
			End If
		End Set
	End Property
	
	Public WriteOnly Property ToDate() As Object
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsDate(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pvToDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mdToDate = Value
			End If
		End Set
	End Property
	
	Public WriteOnly Property RecordSelectionType() As Object
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object pvRecordSelectionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mstrRecordSelectionType = Value
		End Set
	End Property
	
	Public WriteOnly Property AbsencesSelected() As Object
		Set(ByVal Value As Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object pvAbsencesSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mastrSelectedAbsenceTypes = Split(Value, ",")
		End Set
	End Property
	
	Public WriteOnly Property Username() As String
		Set(ByVal Value As String)
			' Username passed in from the asp page
			gsUsername = Value
		End Set
	End Property
	
	Public WriteOnly Property RecordID() As Object
		Set(ByVal Value As Object)
			
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object piRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mlngPersonnelRecordID = Value
			End If
			
		End Set
	End Property
	
	Public WriteOnly Property RealSource() As Object
		Set(ByVal Value As Object)
			
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pstrRealSource. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mstrRealSource = Value
			End If
			
		End Set
	End Property
	
	Public WriteOnly Property IsIndividualRecord() As Object
		Set(ByVal Value As Object)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pbIndividualRecord. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mbIndividualRecord = Value
			
		End Set
	End Property
	
	Public Function HTML_Functions() As Object
		
		Dim iCount As Short
		Dim strAbsenceName As String
		Dim strHTML As String
		Dim strHTML_RunReport As String
		
		' Create function header strings
		strHTML = "<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>" & vbNewLine
		
		strHTML_RunReport = "function RunReport(){" & vbNewLine & "fOK = true;" & "window.parent.frames(""refreshframe"").document.forms(""frmRefresh"").submit();"
		
		' Validate the from date input
		strHTML_RunReport = strHTML_RunReport & vbNewLine & "sValue = txtDateFrom.value;" & vbNewLine & "if (sValue.length == 0) {fOK = false;}" & vbNewLine & " else {" & vbNewLine & "   sValue = convertLocaleDateToSQL(sValue);" & vbNewLine & "     if (sValue.length == 0) {fOK = false;}" & vbNewLine & "     else {txtDateFrom.value = ASRIntranetFunctions.ConvertSQLDateToLocale(sValue);}}" & vbNewLine & "if (fOK == false) {" & vbNewLine & "ASRIntranetFunctions.MessageBox(""Invalid from date value entered."");" & vbNewLine & "txtDateFrom.focus();}"
		
		' Validate the to date input
		strHTML_RunReport = strHTML_RunReport & vbNewLine & "sValue = txtDateTo.value;" & vbNewLine & "if (sValue.length == 0) {fOK = false;}" & vbNewLine & " else {" & vbNewLine & "   sValue = convertLocaleDateToSQL(sValue);" & vbNewLine & "     if (sValue.length == 0) {fOK = false;}" & vbNewLine & "     else {txtDateTo.value = ASRIntranetFunctions.ConvertSQLDateToLocale(sValue);}}" & vbNewLine & "if (fOK == false) {" & vbNewLine & "ASRIntranetFunctions.MessageBox(""Invalid to date value entered."");" & vbNewLine & "txtDateTo.focus();}"
		
		' Pass fields into submit form
		strHTML_RunReport = strHTML_RunReport & vbNewLine & "frmDefinition.txtFromDate.value = txtDateFrom.value;" & vbNewLine & "frmDefinition.txtToDate.value = txtDateTo.value;" & vbNewLine
		
		' Clear the previous options
		strHTML_RunReport = strHTML_RunReport & "frmDefinition.txtAbsenceTypes.value = """"" & ";" & vbNewLine
		
		For iCount = 0 To UBound(mastrAbsenceTypes, 1)
			strAbsenceName = "chkAbsenceType_" & LTrim(Str(iCount))
			strHTML_RunReport = strHTML_RunReport & "if (" & strAbsenceName & ".checked == true) " & "frmDefinition.txtAbsenceTypes.value = frmDefinition.txtAbsenceTypes.value + """ & mastrAbsenceTypes(iCount, 0) & ","";" & vbNewLine
		Next iCount
		
		' Check that an absence type is selected
		strHTML_RunReport = strHTML_RunReport & "if (frmDefinition.txtAbsenceTypes.value == """")" & "{ASRIntranetFunctions.MessageBox(""You must have at least 1 absence type selected."");" & "fOK = false}" & vbNewLine
		
		' Add Bradford Report specific stuff
		If mlngStandardReportType = StandardReportType.utlBradfordFactor Then
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtSRV.value = chkSRV.checked;" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtShowDurations.value = chkShowDurations.checked;" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtShowInstances.value = chkShowInstances.checked;" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtShowFormula.value = chkShowFormula.checked;" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOmitBeforeStart.value = chkOmitBeforeStart.checked;" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOmitAfterEnd.value = chkOmitAfterEnd.checked;" & vbNewLine
			
			If Not mbIndividualRecord Then
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy1.value = cboOrderBy1.value;" & vbNewLine
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy1Asc.value = chkOrderBy1Asc.checked;" & vbNewLine
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy2.value = cboOrderBy2.value;" & vbNewLine
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy2Asc.value = chkOrderBy2Asc.checked;" & vbNewLine
			Else
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy1.value = " & """Surname""" & ";" & vbNewLine
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy1Asc.value = false;" & vbNewLine
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy2.value = " & """Forename""" & ";" & vbNewLine
				strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtOrderBy2Asc.value = false;" & vbNewLine
			End If
			
		End If
		
		' Add in the utility it (filterid or picklistid)
		If Not mbIndividualRecord Then
			strHTML_RunReport = strHTML_RunReport & "if (optPickList.checked == true) frmDefinition.utilid.value = ""0"";" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & "if (optPickList.checked == true) frmDefinition.utilid.value = frmDefinition.txtBasePicklistID.value;" & vbNewLine
			strHTML_RunReport = strHTML_RunReport & "if (optFilter.checked == true) frmDefinition.utilid.value = frmDefinition.txtBaseFilterID.value;" & vbNewLine
		End If
		
		' Check that a picklist is selected
		strHTML_RunReport = strHTML_RunReport & "if ((optPickList.checked == true) && (frmDefinition.txtBasePicklistID.value == ""0"")) " & "{ASRIntranetFunctions.MessageBox(""You must have a picklist selected."");" & "fOK = false}" & vbNewLine
		
		' Check that a filter is selected
		strHTML_RunReport = strHTML_RunReport & "if ((optFilter.checked == true) && (frmDefinition.txtBaseFilterID.value == ""0"")) " & "{ASRIntranetFunctions.MessageBox(""You must have a filter selected."");" & "fOK = false}" & vbNewLine
		
		' Add print base header option
		If Not mbIndividualRecord Then
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtPrintFPinReportHeader.value = chkPrintInReportHeader.checked;" & vbNewLine
		Else
			strHTML_RunReport = strHTML_RunReport & " frmDefinition.txtPrintFPinReportHeader.value = false" & vbNewLine
		End If
		
		' Do the submit
		strHTML_RunReport = strHTML_RunReport & "if (fOK == true) {" & vbNewLine & "var sUtilID = new String(" & Trim(Str(mlngStandardReportType)) & ");" & vbNewLine & "frmDefinition.target = sUtilID;" & vbNewLine & "openWindow('', sUtilID ,'500','200','yes');" & vbNewLine & "frmDefinition.submit();}" & vbNewLine & "return false;}"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_Functions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_Functions = strHTML & strHTML_RunReport & vbNewLine & "</SCRIPT>" & vbNewLine
		
	End Function
	
	
	Public Function HTML_RecordSelection() As Object
		
		Dim strHTML As String
		Dim strAbsenceName As String
		Dim iCount As Short
		Dim dFromDate As Date
		Dim dToDate As Date
		
		dToDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, VB.Day(Today) * -1, Today)
		dFromDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, dToDate))
		
		' Identify this as div1
		strHTML = "<TR><TD><DIV id=div1><TABLE width=100% border=1 cellspacing=0 cellpadding=5>" & "<TR><TD>" & "<TABLE width=100% border=1 cellspacing=0 cellpadding=5>" & "<TD bgcolor=threedface><TABLE width=""100%"" border=0 CELLSPACING=0 CELLPADDING=0>"
		
		' Define date range
		strHTML = strHTML & "<TR bordercolor=" & mstrHexColour_FormBackground & ">" & "<TD>Date Range</TD>" & "<TD><INPUT id=txtDateFrom name=txtDateFrom value=" & VB6.Format(dFromDate, "dd/mm/yyyy") & "></TD>" & "<TD><INPUT id=txtDateTo name=txtDateTo value=" & VB6.Format(dToDate, "dd/mm/yyyy") & "></TD></TR>" & vbNewLine
		
		' Add a blank line
		strHTML = strHTML & "<TR><TD colspan=3>&nbsp</TD></TR>"
		
		' Define absence types
		strHTML = strHTML & "<TR bordercolor=" & mstrHexColour_FormBackground & ">" & "<TD>Absence Types</TD>" & "<TD colspan=2>" & "<SPAN ID=spanAbsenceTypes STYLE=""width:300px;height:120px; overflow:auto"">" & "<TABLE cellspacing=0 cellpadding=0 bgColor=" & mstrHexColour_AbsenceTypesBackground & " width=100%>" & vbNewLine
		
		' Stick in the absence types
		For iCount = 0 To UBound(mastrAbsenceTypes, 1)
			strAbsenceName = "chkAbsenceType_" & LTrim(Str(iCount))
			strHTML = strHTML & "<TR><TD>" & "<INPUT id=" & strAbsenceName & " name=" & strAbsenceName & " type=checkbox " & IIf(CBool(mastrAbsenceTypes(iCount, 2)), "CHECKED", "") & ">" & mastrAbsenceTypes(iCount, 0) & "</TD></TR>" & vbNewLine
		Next iCount
		
		' Round off top half of options
		strHTML = strHTML & "<TR><TD colspan=3><HR></TD></TR></SPAN></TD></TR></TABLE></TABLE>"
		
		' Add a blank line
		strHTML = strHTML & "<TR><TABLE border=0 cellspacing=0 cellpadding=3>&nbsp</TR>"
		
		' Add the all/picklist/filter table
		If Not mbIndividualRecord Then
			
			strHTML = strHTML & "<TR>" & "<TABLE WIDTH=""100%"" height=""80%"" border=1 cellspacing=0 cellpadding=5>" & "<TD bgcolor=threedface><TABLE WIDTH=""100%"" border=0 CELLSPACING=0 CELLPADDING=0>"
			
			' Generate the record selection (All)
			strHTML = strHTML & "<TR bordercolor=" & mstrHexColour_FormBackground & ">" & "<TD width=75 colspan=3>" & "<INPUT CHECKED id=optAllRecords name=optAllRecords type=radio onclick=""changeRecordOptions('all')"">" & "All</TD></TR>"
			
			' Generate picklist record selection
			strHTML = strHTML & "<TR bordercolor=" & mstrHexColour_FormBackground & ">" & "<TD width=75>" & "<INPUT id=optPickList name=optPickList type=radio onclick=""changeRecordOptions('picklist')"">" & "Picklist</TD><TD>" & "<INPUT id=txtBasePicklist name=txtBasePicklist readOnly style=""BACKGROUND-COLOR: ThreeDFace;WIDTH:100%""></TD>" & "<TD width=15>" & "<INPUT id=cmdBasePicklist name=cmdBasePicklist disabled type=button value=... onclick=""selectRecordOption('picklist')"" onfocus=""refreshControls()"" > " & "</TD></TR>"
			
			' Generate filter record selection
			strHTML = strHTML & "<TR bordercolor=" & mstrHexColour_FormBackground & "><TD>" & "<INPUT id=optFilter name=optFilter type=radio onclick=""changeRecordOptions('filter')"">" & "Filter</TD><TD>" & "<INPUT id=txtBaseFilter name=txtBaseFilter readOnly style=""BACKGROUND-COLOR: ThreeDFace;WIDTH:100%""></TD>" & "<TD><INPUT id=cmdBaseFilter name=cmdBaseFilter disabled type=button onclick=""selectRecordOption('filter')"" onfocus=refreshControls() value=...>" & "</TD></TR></TD></TR></TABLE></TABLE>"
			
		End If
		
		' Generate the print report options
		If Not mbIndividualRecord Then
			strHTML = strHTML & "<P><TABLE WIDTH=""100%"" height=""80%"" border=1 cellspacing=0 cellpadding=5>" & "<TR><TD><INPUT id=chkPrintIneportHeader name=chkPrintInReportHeader type=checkbox>" & "Print Base Table filter/picklist in the report header</TD></TR></TABLE>"
		Else
			strHTML = strHTML & "<TABLE></TABLE>"
		End If
		
		
		'End this division
		strHTML = strHTML & "</TABLE></DIV>"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_RecordSelection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_RecordSelection = strHTML
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		' Define colours for this form
		mstrHexColour_FormBackground = "silver"
		mstrHexColour_FormText = " white"
		mstrHexColour_AbsenceTypesBackground = "white"
		
		mlngStandardReportType = StandardReportType.utlUndefined
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub LoadOptions()
		
		Dim bOK As Boolean
		
		On Error GoTo AlreadyLoaded
		
		' Load basic stuff
		SetupTablesCollection()
		
		' Read the personnel settings
		'If glngPersonnelTableID = 0 Then
		ReadPersonnelParameters()
		'End If
		
		' Load the absence settings
		'If glngAbsenceTableID = 0 Then
		ReadAbsenceParameters()
		'End If
		
LoadAbsenceTypes: 
		bOK = LoadAbsenceTypes
		Exit Sub
		
AlreadyLoaded: 
		bOK = False
		
	End Sub
	
	' Load the absence types from the system
	Private Function LoadAbsenceTypes() As Boolean
		
		On Error GoTo errLoadAbsenceTypes
		
		Dim rstAbsenceTypes As ADODB.Recordset
		Dim iCount As Short
		Dim strSQL As String
		Dim strIncludeColumnName As String
		
		' What include column should we be looking at
		Select Case mlngStandardReportType
			
			Case StandardReportType.utlAbsenceBreakdown
				strIncludeColumnName = gsAbsenceTypeIncludeColumnName
				
			Case StandardReportType.utlBradfordFactor
				strIncludeColumnName = gsAbsenceTypeBradfordIndexColumnName
				
		End Select
		
		' Build string to suck in absence types
		strSQL = "SELECT DISTINCT " & gsAbsenceTypeTypeColumnName & " AS Type," & gsAbsenceTypeCodeColumnName & " AS TypeCode" & IIf(Len(strIncludeColumnName) > 0, ", " & strIncludeColumnName & " AS Include", "") & " FROM " & gsAbsenceTypeTableName & " ORDER BY " & gsAbsenceTypeTypeColumnName
		rstAbsenceTypes = datGeneral.GetRecords(strSQL)
		
		If rstAbsenceTypes.BOF And rstAbsenceTypes.EOF Then
			LoadAbsenceTypes = False
			Exit Function
		End If
		
		' Populate absence type array
		ReDim mastrAbsenceTypes(rstAbsenceTypes.RecordCount - 1, 2)
		iCount = 0
		Do Until rstAbsenceTypes.EOF
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mastrAbsenceTypes(iCount, 0) = IIf(IsDbNull(rstAbsenceTypes.Fields("Type").Value), "", rstAbsenceTypes.Fields("Type").Value)
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			mastrAbsenceTypes(iCount, 1) = IIf(IsDbNull(rstAbsenceTypes.Fields("TypeCode").Value), "", rstAbsenceTypes.Fields("TypeCode").Value)
			
			If Len(strIncludeColumnName) > 0 Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				mastrAbsenceTypes(iCount, 2) = IIf(IsDbNull(rstAbsenceTypes.Fields("Include").Value), False, rstAbsenceTypes.Fields("Include").Value)
			Else
				mastrAbsenceTypes(iCount, 2) = CStr(False)
			End If
			
			rstAbsenceTypes.MoveNext()
			iCount = iCount + 1
		Loop 
		
		' If we are here, then notify calling procedure of success and exit
		'UPGRADE_NOTE: Object rstAbsenceTypes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAbsenceTypes = Nothing
		LoadAbsenceTypes = True
		Exit Function
		
errLoadAbsenceTypes: 
		
		'UPGRADE_NOTE: Object rstAbsenceTypes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAbsenceTypes = Nothing
		LoadAbsenceTypes = False
		
	End Function
	
	Public Function HTML_Title() As Object
		
		Dim strHTML As String
		Dim strCaption As String
		
		Select Case mlngStandardReportType
			
			Case StandardReportType.utlAbsenceBreakdown
				strCaption = "Absence Breakdown"
				
			Case StandardReportType.utlBradfordFactor
				strCaption = "Bradford Factor"
				
		End Select
		
		strHTML = "<TR><TD align=center colspan=3><H2>" & strCaption & "</H2></TD></TR>"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_Title. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_Title = strHTML
		
	End Function
	
	Public Function HTML_BradfordOptions() As Object
		
		Dim strHTML As String
		Dim iCount As Short
		
		' Identify this as div1
		strHTML = "<DIV id=div2 style=""display:none""><TABLE width=100% border=1 cellspacing=0 cellpadding=5>"
		
		Select Case mlngStandardReportType
			
			Case StandardReportType.utlAbsenceBreakdown
				strHTML = strHTML
				
			Case StandardReportType.utlBradfordFactor
				
				' Load the personnel field list
				GetPersonnelFieldList()
				
				' Add display options
				strHTML = strHTML & "<TR><TD><TABLE width=100% border=1 cellspacing=0 cellpadding=5><TD>" & "<TABLE border=0 cellspacing=0 cellpadding=0></TD></TR>" & "<TR><TD><INPUT type=""checkbox"" id=chkSRV name=chkSRV>Suppress Repeated Personnel Details</TD></TR>" & "<TR><TD><INPUT type=""checkbox"" checked id=chkShowDurations name=chkShowDurations>Show Duration Totals</TD></TR>" & "<TR><TD><INPUT type=""checkbox"" id=chkShowInstances name=chkShowInstances>Show Instances Count</TD></TR>" & "<TR><TD><INPUT type=""checkbox"" id=chkShowFormula name=chkShowFormula>Show Bradford Factor Formula</TD></TR>" & "</TABLE></TABLE><P>"
				
				' Add the included record options
				strHTML = strHTML & "<TABLE width=100% border=1 cellspacing=0 cellpadding=5><TD>" & "<TABLE border=0 cellspacing=0 cellpadding=0></TD></TR>" & "<TR><TD><INPUT type=""checkbox"" id=chkOmitBeforeStart name=chkOmitBeforeStart>Omit absences starting before the report period</TD></TR>" & "<TR><TD><INPUT type=""checkbox"" id=chkOmitAfterEnd name=chkOmitAfterEnd>Omit absences ending after the report period</TD></TR>" & "</TABLE></TABLE><P>"
				
				' Add the order by and then by
				If Not mbIndividualRecord Then
					
					strHTML = strHTML & "<TABLE width=100% border=1 cellspacing=0 cellpadding=5><TD>" & "<TABLE width=100% border=0 cellspacing=0 cellpadding=0>" & "<TR><TD>Order By</TD><TD width=50%><SELECT id=cboOrderBy1 name=cboOrderBy1 style=""WIDTH: 50%"">"
					
					For iCount = 0 To UBound(mastrPersonnelFieldList) - 1
						If gsPersonnelSurnameColumnName = mastrPersonnelFieldList(iCount) Then
							strHTML = strHTML & "<OPTION selected VALUE=""" & mastrPersonnelFieldList(iCount) & """>" & mastrPersonnelFieldList(iCount) & "</OPTION>"
						Else
							strHTML = strHTML & "<OPTION VALUE=""" & mastrPersonnelFieldList(iCount) & """>" & mastrPersonnelFieldList(iCount) & "</OPTION>"
						End If
					Next iCount
					
					strHTML = strHTML & "</SELECT></TD><TD><INPUT type=""checkbox"" checked id=chkOrderBy1Asc name=chkOrderBy1Asc>" & "Ascending</TD></TR>" & "<TR><TD>Then By</TD><TD width=50%><SELECT id=cboOrderBy2 name=cboOrderBy2 style=""WIDTH: 50%"">"
					
					For iCount = 0 To UBound(mastrPersonnelFieldList) - 1
						If gsPersonnelForenameColumnName = mastrPersonnelFieldList(iCount) Then
							strHTML = strHTML & "<OPTION selected VALUE=""" & mastrPersonnelFieldList(iCount) & """>" & mastrPersonnelFieldList(iCount) & "</OPTION>"
						Else
							strHTML = strHTML & "<OPTION VALUE=""" & mastrPersonnelFieldList(iCount) & """>" & mastrPersonnelFieldList(iCount) & "</OPTION>"
						End If
					Next iCount
					
					strHTML = strHTML & "</SELECT></TD><TD><INPUT type=""checkbox"" checked id=chkOrderBy2Asc name=chkOrderBy2Asc>" & "Ascending</TD></TR></TABLE></TABLE>"
					
				End If
				
			Case Else
				strHTML = strHTML
				
		End Select
		
		'End this division
		strHTML = strHTML & "</TABLE></DIV>"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_BradfordOptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_BradfordOptions = strHTML
		
	End Function
	
	Public Function HTML_PageButtons() As Object
		
		Dim strHTML As String
		
		Select Case mlngStandardReportType
			
			Case StandardReportType.utlAbsenceBreakdown
				strHTML = ""
				
			Case StandardReportType.utlBradfordFactor
				
				strHTML = "<TR><TD colspan=3>" & "<table width=100% align=left border=0 cellPadding=5 cellSpacing=0>" & "<tr height=10>" & "<td>" & "<INPUT type=""button"" value=""Selection"" id=btnTab1 name=btnTab1 onclick=""displayPage(1)"" disabled=true>" & "<INPUT type=""button"" value=""Options"" id=btnTab2 name=btnTab2 onclick=""displayPage(2)"">" & "</td></tr>" & "</TD></TR>"
				
			Case Else
				strHTML = ""
				
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_PageButtons. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_PageButtons = strHTML
		
	End Function
	
	
	Private Sub GetPersonnelFieldList()
		
		Dim fOK As Boolean
		On Error GoTo ErrorTrap
		
		' Loads the Base combo with all columns for personnel records
		
		Dim sSQL As String
		Dim rsColumns As New ADODB.Recordset
		Dim datData As New clsDataAccess
		Dim iCount As Short
		
		sSQL = "Select ColumnID, ColumnName From ASRSysColumns WHERE TableID = " & glngPersonnelTableID
		sSQL = sSQL & " ORDER BY ColumnName"
		rsColumns = datGeneral.GetRecords(sSQL)
		
		If Not rsColumns.EOF And Not rsColumns.BOF Then
			ReDim mastrPersonnelFieldList(rsColumns.RecordCount)
			
			'Add blanks
			mastrPersonnelFieldList(0) = "None"
			
			' Populate with columns from personnel records
			rsColumns.MoveFirst()
			iCount = 1
			Do While Not rsColumns.EOF
				
				If Not rsColumns.Fields("ColumnName").Value = "ID" Then
					mastrPersonnelFieldList(iCount) = rsColumns.Fields("ColumnName").Value
					iCount = iCount + 1
				End If
				
				rsColumns.MoveNext()
			Loop 
			
			rsColumns.Close()
			'UPGRADE_NOTE: Object rsColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsColumns = Nothing
			
			ReDim Preserve mastrPersonnelFieldList(iCount)
			
		Else
			ReDim mastrPersonnelFieldList(1)
			mastrPersonnelFieldList(0) = "<None>"
			fOK = False
		End If
		
TidyUpAndExit: 
		Exit Sub
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Sub
End Class