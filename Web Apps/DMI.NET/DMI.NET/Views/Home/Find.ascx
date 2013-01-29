<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%
	On Error Resume Next

	'Dim sReferringPage

	
	
	'' Only open the form if there was a referring page.
	'' If it wasn't then redirect to the login page.
	'sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	'if inStrRev(sReferringPage, "/") > 0 then
	'	sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	'end if

	'if len(sReferringPage) = 0 then
	'	Response.Redirect("login.asp")
	'end if
	

	' Flag an error if there is no current table or view is specified.
	If (Session("tableID") <= 0) And _
	 (Session("viewID") <= 0) Then
	
		Session("ErrorTitle") = "Find Page"
		Session("ErrorText") = "No table or view specified."
		Response.Redirect("error")
	End If
	
	' Flag an error if there is no current screen is specified.
	If Session("screenID") <= 0 Then
		Session("ErrorTitle") = "Find Page"
		Session("ErrorText") = "No screen specified."
		Response.Redirect("error")
	End If
	
	' Get the screen's default order if none is already specified.
	If Session("orderID") <= 0 Then
		Dim cmdScreenOrder = CreateObject("ADODB.Command")
		cmdScreenOrder.CommandText = "sp_ASRIntGetScreenOrder"
		cmdScreenOrder.CommandType = 4 ' Stored Procedure
		cmdScreenOrder.ActiveConnection = Session("databaseConnection")

		Dim prmOrderID = cmdScreenOrder.CreateParameter("orderID", 3, 2)
		cmdScreenOrder.Parameters.Append(prmOrderID)

		Dim prmScreenID2 = cmdScreenOrder.CreateParameter("screenID", 3, 1)
		cmdScreenOrder.Parameters.Append(prmScreenID2)
		prmScreenID2.value = CleanNumeric(Session("screenID"))

		Err.Clear()
		cmdScreenOrder.Execute()
		If (Err.Number <> 0) Then
			Session("ErrorTitle") = "Find Page"
			Session("ErrorText") = "The default order for the screen could not be determined :<p>" & formatError(Err.Description)
			Response.Redirect("error")
		Else
			Session("orderID") = cmdScreenOrder.Parameters("orderID").Value
		End If
		' Release the ADO command object.
		cmdScreenOrder = Nothing
	End If

	' Enable response buffering as we may redirect the response further down this page.
	Response.Buffer = True
%>
	<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0"
		viewastext>
		<param name="LPKPath" value="lpks/main.lpk">
	</object>
	<script type="text/javascript">
<!--
		
		
		function find_window_onload() {
			var fOk;

			fOk = true;

			var frmFindForm = document.getElementById("frmFindForm");

			var sErrMsg = frmFindForm.txtErrorDescription.value;
			if (sErrMsg.length > 0) {
				fOk = false;

//				window.parent.frames("menuframe").ASRIntranetFunctions.Closepopup();
				OpenHR.messageBox(sErrMsg);
				menu_loadPage("default");				
			}

			if (fOk == true) {

				setGridFont(frmFindForm.ssOleDBGridFindRecords);

				var frmMenuInfo = document.getElementById("frmMenuInfo");

				if ((frmMenuInfo.txtUserType.value == 1) &&
					(frmMenuInfo.txtPersonnel_EmpTableID.value == frmFindForm.txtCurrentTableID.value) &&
					(frmFindForm.txtRecordCount.value > 1)) {

					frmFindForm.ssOleDBGridFindRecords.focus();
					frmFindForm.ssOleDBGridFindRecords.RemoveAll();

					// Get menu.asp to refresh the menu.
					menu_refreshMenu();

					/* The user does NOT have permission to create new records. */
					OpenHR.messageBox("Unable to load personnel records.\n\nYou are logged on as a self-service user and can access only single record personnel record sets.");

					/* Go to the default page. */
					menu_loadPage("default");
					return;
				}
			}

			if (fOk == true) {
				var sControlName;
				var sControlPrefix;
				var sColumnId;
				var ctlSummaryControl;
				var sSummaryControlName;
				var sDataType;

				// Expand the work frame and hide the option frame.
				//window.parent.document.all.item("workframeset").cols = "*, 0";
				$("#workframe").attr("data-framesource", "FIND");

				// JPD20020903 Fault 2316 - Need to dim focus on the grid before adding the items.
				frmFindForm.ssOleDBGridFindRecords.focus();
			    
				var controlCollection = frmFindForm.elements;
				if (controlCollection != null) {
					for (var i = 0; i < controlCollection.length; i++) {

						sControlName = controlCollection.item(i).name;
						sControlPrefix = sControlName.substr(0, 13);

						if (sControlPrefix == "txtAddString_") {
							frmFindForm.ssOleDBGridFindRecords.AddItem(controlCollection.item(i).value);
						}

						sControlName = controlCollection.item(i).name;
						sControlPrefix = sControlName.substr(0, 15);

						if (sControlPrefix == "txtSummaryData_") {
							sColumnId = sControlName.substr(15);
							sSummaryControlName = "ctlSummary_";
							sSummaryControlName = sSummaryControlName.concat(sColumnId);
							sSummaryControlName = sSummaryControlName.concat("_");

							for (var j = 0; j < controlCollection.length; j++) {
								sControlName = controlCollection.item(j).name;
								sControlPrefix = sControlName.substr(0, sSummaryControlName.length);

								if (sControlPrefix == sSummaryControlName) {
									ctlSummaryControl = controlCollection.item(j);

									if (ctlSummaryControl.type == "checkbox") {
										ctlSummaryControl.checked = (controlCollection.item(i).value.toUpperCase() == "TRUE");
									} else {
										// Check if the control is for a datevalue.
										sDataType = sControlName.substr(sSummaryControlName.length);

										if (sDataType == "11") {
											// Format dates for the locale setting.							
											if (controlCollection.item(i).value == '') {
												ctlSummaryControl.value = '';
											} else {
												//TODO:ctlSummaryControl.value = window.parent.frames("menuframe").ASRIntranetFunctions.ConvertSQLDateToLocale(controlCollection.item(i).value);
												ctlSummaryControl.value = controlCollection.item(i).value;
											}
										} else {
											ctlSummaryControl.value = controlCollection.item(i).value;
										}
									}

									break;
								}
							}
						}
					}
				}

				// dim focus onto one of the form controls. 
				// NB. This needs to be done before making any reference to the grid
				frmFindForm.ssOleDBGridFindRecords.focus();


				// Select the current record in the grid if its there, else select the top record if there is one.
				if (frmFindForm.ssOleDBGridFindRecords.rows > 0) {
					if ((frmFindForm.txtCurrentRecordID.value > 0) && (frmFindForm.txtGotoAction.value != 'LOCATE')) {
						// Try to select the current record.
						locateRecord(frmFindForm.txtCurrentRecordID.value, true);
					} else {
						// Select the top row.
						frmFindForm.ssOleDBGridFindRecords.MoveFirst();
						frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
					}
				}

				// Get menu.asp to refresh the menu.
				menu_refreshMenu();

				if ((frmFindForm.ssOleDBGridFindRecords.rows == 0) && (frmFindForm.txtFilterSQL.value.length > 0)) {
					OpenHR.messageBox("No records match the current filter.\nNo filter is applied.");
					menu_clearFilter();
				}
			}
		}
	-->
	</script>
	<script type="text/javascript">
<!--
		/* Return the ID of the record selected in the find form. */
	    function selectedRecordID() {
	        
			var iRecordId;
			var iIndex;
			var iIdColumnIndex;
			var sColumnName;

			var frmFindForm = document.getElementById("frmFindForm");

			iRecordId = 0;
			iIdColumnIndex = 0;

			if (frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Count > 0) {
				for (iIndex = 0; iIndex < frmFindForm.ssOleDBGridFindRecords.Cols; iIndex++) {
					sColumnName = frmFindForm.ssOleDBGridFindRecords.Columns(iIndex).Name;

					if (sColumnName.toUpperCase() == "ID") {
						iIdColumnIndex = iIndex;
						break;
					}
				}

				iRecordId = frmFindForm.ssOleDBGridFindRecords.Columns(iIdColumnIndex).Value;

			}

			return (iRecordId);
		}

		/* Sequential search the grid for the required ID. */
		function locateRecord(psSearchFor, pfIdMatch) {
			var fFound;
			var iIndex;
			var iIdColumnIndex;
			var sColumnName;

			var frmFindForm = document.getElementById("frmFindForm");

			fFound = false;
			frmFindForm.ssOleDBGridFindRecords.redraw = false;

			if (pfIdMatch == true) {
				// Locate the ID column in the grid.
				iIdColumnIndex = -1;
				for (iIndex = 0; iIndex < frmFindForm.ssOleDBGridFindRecords.Cols; iIndex++) {
					sColumnName = frmFindForm.ssOleDBGridFindRecords.Columns(iIndex).Name;
					if (sColumnName.toUpperCase() == "ID") {
						iIdColumnIndex = iIndex;
						break;
					}
				}

				if (iIdColumnIndex >= 0) {
					frmFindForm.ssOleDBGridFindRecords.MoveLast();
					frmFindForm.ssOleDBGridFindRecords.MoveFirst();

					for (iIndex = 1; iIndex <= frmFindForm.ssOleDBGridFindRecords.rows; iIndex++) {
						if (frmFindForm.ssOleDBGridFindRecords.Columns(iIdColumnIndex).value == psSearchFor) {
							frmFindForm.ssOleDBGridFindRecords.FirstRow = frmFindForm.ssOleDBGridFindRecords.Bookmark;
							if ((frmFindForm.ssOleDBGridFindRecords.Rows - frmFindForm.ssOleDBGridFindRecords.AddItemRowIndex(frmFindForm.ssOleDBGridFindRecords.FirstRow) + 1) < frmFindForm.ssOleDBGridFindRecords.VisibleRows) {
								if (frmFindForm.ssOleDBGridFindRecords.Rows - frmFindForm.ssOleDBGridFindRecords.VisibleRows + 1 >= 1) {
									frmFindForm.ssOleDBGridFindRecords.FirstRow = frmFindForm.ssOleDBGridFindRecords.AddItemBookmark(frmFindForm.ssOleDBGridFindRecords.Rows - frmFindForm.ssOleDBGridFindRecords.VisibleRows + 1);
								}
								else {
									frmFindForm.ssOleDBGridFindRecords.FirstRow = frmFindForm.ssOleDBGridFindRecords.AddItemBookmark(0);
								}
							}

							frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
							fFound = true;
							break;
						}

						if (iIndex < frmFindForm.ssOleDBGridFindRecords.rows) {
							frmFindForm.ssOleDBGridFindRecords.MoveNext();
						}
						else {
							break;
						}
					}
				}
			}
			else {
				for (iIndex = 1; iIndex <= frmFindForm.ssOleDBGridFindRecords.rows; iIndex++) {
					var sGridValue = new String(frmFindForm.ssOleDBGridFindRecords.Columns(0).value);
					sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
					if (sGridValue == psSearchFor.toUpperCase()) {
						frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
						fFound = true;
						break;
					}

					if (iIndex < frmFindForm.ssOleDBGridFindRecords.rows) {
						frmFindForm.ssOleDBGridFindRecords.MoveNext();
					}
					else {
						break;
					}
				}
			}

			if ((fFound == false) && (frmFindForm.ssOleDBGridFindRecords.rows > 0)) {
				// Select the top row.
				frmFindForm.ssOleDBGridFindRecords.MoveFirst();
				frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
			}

			frmFindForm.ssOleDBGridFindRecords.redraw = true;
		}
	-->
	</script>
	<script type="text/javascript">		
		<!--		
		// Double-click in the grid. Edit the selected record.
		OpenHR.addActiveXHandler("ssOleDBGridFindRecords", "dblClick", ssOleDBGridFindRecords_dblClick);

		function ssOleDBGridFindRecords_dblClick() {
			menu_editRecord();
		}
	-->
	</script>
	<script type="text/javascript">
<!--
		OpenHR.addActiveXHandler("ssOleDBGridFindRecords", "KeyPress", ssOleDBGridFindRecords_KeyPress);

		function ssOleDBGridFindRecords_KeyPress(iKeyAscii) {
			if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
				var frmFindForm = document.getElementById("frmFindForm");
				var dtTicker = new Date();
				var iThisTick = new Number(dtTicker.getTime());
				var iLastTick;
				
				if (frmFindForm.txtLastKeyFind.value.length > 0) {
					iLastTick = new Number(frmFindForm.txtTicker.value);
				} else {
					iLastTick = new Number("0");
				}
				var sFind;
				if (iThisTick > (iLastTick + 1500)) {
					sFind = String.fromCharCode(iKeyAscii);
				} else {
					sFind = frmFindForm.txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
				}

				frmFindForm.txtTicker.value = iThisTick;
				frmFindForm.txtLastKeyFind.value = sFind;

				locateRecord(sFind, false);
			}
		}
	-->
	</script>
	<script type="text/javascript">
<!--
		OpenHR.addActiveXHandler("ssOleDBGridFindRecords", "click", ssOleDBGridFindRecords_Click);

		function ssOleDBGridFindRecords_Click() {
			// Click in the grid. Refresh the menu.		
			menu_refreshMenu();
		}
	-->
	</script>
<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmFindForm" name="frmFindForm">
	<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%"
		height="100%">
		<tr>
			<td>
				<table id="findTable" width="100%" height="100%" class="invisible" cellspacing="0"
					cellpadding="0">
					<tr height="10">
						<td align="center" colspan="5" height="10">
							<%
								On Error Resume Next
	
								Dim sErrorDescription As String = ""
	
								' Display the appropriate page title.
								Dim cmdFindWindowTitle = CreateObject("ADODB.Command")
								cmdFindWindowTitle.CommandText = "sp_ASRIntGetFindWindowInfo"
								cmdFindWindowTitle.CommandType = 4 ' Stored Procedure
								cmdFindWindowTitle.ActiveConnection = Session("databaseConnection")

								Dim prmTitle = cmdFindWindowTitle.CreateParameter("title", 200, 2, 100)
								cmdFindWindowTitle.Parameters.Append(prmTitle)

								Dim prmQuickEntry = cmdFindWindowTitle.CreateParameter("quickEntry", 11, 2) ' 11=bit, 2=output
								cmdFindWindowTitle.Parameters.Append(prmQuickEntry)

								Dim prmScreenID = cmdFindWindowTitle.CreateParameter("screenID", 3, 1)
								cmdFindWindowTitle.Parameters.Append(prmScreenID)
								prmScreenID.value = CleanNumeric(Session("screenID"))

								Dim prmViewID = cmdFindWindowTitle.CreateParameter("viewID", 3, 1)
								cmdFindWindowTitle.Parameters.Append(prmViewID)
								prmViewID.value = CleanNumeric(Session("viewID"))

								Err.Clear()
								cmdFindWindowTitle.Execute()
								If (Err.Number <> 0) Then
									sErrorDescription = "The page title could not be created." & vbCrLf & formatError(Err.Description)
								End If

								If Len(sErrorDescription) = 0 Then
									Response.Write("						<h3 align=center>Find - " & Replace(cmdFindWindowTitle.Parameters("title").Value, "_", " ") & "</h3>" & vbCrLf)
									Response.Write("<INPUT type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & cmdFindWindowTitle.Parameters("quickEntry").Value & ">" & vbCrLf)
								End If
	
								' Release the ADO command object.
								cmdFindWindowTitle = Nothing
							%>
						</td>
					</tr>
					<tr id="findGridRow">
						<td>
						</td>
						<td width="100%" colspan="3" height="500">
							<%
								Dim sThousandColumns As String
								Dim sBlankIfZeroColumns As String
								Dim sTemp As String
								
								If Len(sErrorDescription) = 0 Then									
									' Get the find records.
									Dim cmdThousandFindColumns = CreateObject("ADODB.Command")
									cmdThousandFindColumns.CommandText = "spASRIntGet1000SeparatorFindColumns"
									cmdThousandFindColumns.CommandType = 4	' Stored Procedure
									cmdThousandFindColumns.ActiveConnection = Session("databaseConnection")
									cmdThousandFindColumns.CommandTimeout = 180
		
									Dim prmError = cmdThousandFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
									cmdThousandFindColumns.Parameters.Append(prmError)

									Dim prmTableID = cmdThousandFindColumns.CreateParameter("tableID", 3, 1)
									cmdThousandFindColumns.Parameters.Append(prmTableID)
									prmTableID.value = CleanNumeric(Session("tableID"))

									prmViewID = cmdThousandFindColumns.CreateParameter("viewID", 3, 1)
									cmdThousandFindColumns.Parameters.Append(prmViewID)
									prmViewID.value = CleanNumeric(Session("viewID"))

									Dim prmOrderID = cmdThousandFindColumns.CreateParameter("orderID", 3, 1)
									cmdThousandFindColumns.Parameters.Append(prmOrderID)
									prmOrderID.value = CleanNumeric(Session("orderID"))

									Dim prmThousandColumns = cmdThousandFindColumns.CreateParameter("thousandColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
									cmdThousandFindColumns.Parameters.Append(prmThousandColumns)
	
									Err.Clear()
									cmdThousandFindColumns.Execute()

									If (Err.Number <> 0) Then
										sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
									End If

									If Len(sErrorDescription) = 0 Then
										sThousandColumns = cmdThousandFindColumns.Parameters("thousandColumns").Value
									End If
	
									' Release the ADO command object.
									cmdThousandFindColumns = Nothing
								End If

								' NPG20090210 Fault 13249
								If Len(sErrorDescription) = 0 Then
									' Get the BlankIfZero find records.
									Dim cmdBlankIfZeroFindColumns = CreateObject("ADODB.Command")
									cmdBlankIfZeroFindColumns.CommandText = "spASRIntGetBlankIfZeroFindColumns"
									cmdBlankIfZeroFindColumns.CommandType = 4	' Stored Procedure
									cmdBlankIfZeroFindColumns.ActiveConnection = Session("databaseConnection")
									cmdBlankIfZeroFindColumns.CommandTimeout = 180
		
									Dim prmError = cmdBlankIfZeroFindColumns.CreateParameter("error", 11, 2) ' 11=bit, 2=output
									cmdBlankIfZeroFindColumns.Parameters.Append(prmError)

									Dim prmTableID = cmdBlankIfZeroFindColumns.CreateParameter("tableID", 3, 1)
									cmdBlankIfZeroFindColumns.Parameters.Append(prmTableID)
									prmTableID.value = CleanNumeric(CLng(Session("tableID")))

									prmViewID = cmdBlankIfZeroFindColumns.CreateParameter("viewID", 3, 1)
									cmdBlankIfZeroFindColumns.Parameters.Append(prmViewID)
									prmViewID.value = CleanNumeric(CLng(Session("viewID")))

									Dim prmOrderID = cmdBlankIfZeroFindColumns.CreateParameter("orderID", 3, 1)
									cmdBlankIfZeroFindColumns.Parameters.Append(prmOrderID)
									prmOrderID.value = CleanNumeric(CLng(Session("orderID")))

									Dim prmBlankIfZeroColumns = cmdBlankIfZeroFindColumns.CreateParameter("blankifzeroColumns", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
									cmdBlankIfZeroFindColumns.Parameters.Append(prmBlankIfZeroColumns)
	
									Err.Clear()
									cmdBlankIfZeroFindColumns.Execute()

									If (Err.Number <> 0) Then
										sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
									End If

									If Len(sErrorDescription) = 0 Then
										sBlankIfZeroColumns = cmdBlankIfZeroFindColumns.Parameters("blankifzeroColumns").Value
									End If
	
									' Release the ADO command object.
									cmdBlankIfZeroFindColumns = Nothing
								End If

								Dim fCancelDateColumn = True
								If (Len(sErrorDescription) = 0) And (Session("TB_CourseTableID") > 0) Then
									Dim sSubString As String = Session("lineage")
									Dim iIndex = InStr(sSubString, "_")
									sSubString = Mid(sSubString, iIndex + 1)
									iIndex = InStr(sSubString, "_")
									sSubString = Mid(sSubString, iIndex + 1)
									iIndex = InStr(sSubString, "_")
									sSubString = Mid(sSubString, iIndex + 1)
									iIndex = InStr(sSubString, "_")
									sSubString = Mid(sSubString, iIndex + 1)
									iIndex = InStr(sSubString, "_")
									Dim lngRecordID = Left(sSubString, iIndex - 1)

									' Get the Course Date
									Dim cmdGetCancelDateColumn = CreateObject("ADODB.Command")
									cmdGetCancelDateColumn.CommandText = "spASRIntGetCancelCourseDate"
									cmdGetCancelDateColumn.CommandType = 4	' Stored Procedure
									cmdGetCancelDateColumn.ActiveConnection = Session("databaseConnection")
									cmdGetCancelDateColumn.CommandTimeout = 180
				
									Dim prmError = cmdGetCancelDateColumn.CreateParameter("error", 11, 2) ' 11=bit, 2=output
									cmdGetCancelDateColumn.Parameters.Append(prmError)

									Dim prmRecID = cmdGetCancelDateColumn.CreateParameter("recordID", 3, 1)	' 3=integer, 1=input
									cmdGetCancelDateColumn.Parameters.Append(prmRecID)
									prmRecID.value = CleanNumeric(lngRecordID)

									Dim prmCancelDateColumn = cmdGetCancelDateColumn.CreateParameter("CancelDateColumn", 11, 2) ' 11=bit, 2=output
									cmdGetCancelDateColumn.Parameters.Append(prmCancelDateColumn)
			
									Err.Clear()
									cmdGetCancelDateColumn.Execute()

									If (Err.Number <> 0) Then
										sErrorDescription = "Unable to check for a Cancelled Course Date." & vbCrLf & formatError(Err.Description)
									End If

									If Len(sErrorDescription) = 0 Then
										fCancelDateColumn = cmdGetCancelDateColumn.Parameters("CancelDateColumn").Value
									End If
			
									' Release the ADO command object.
									cmdGetCancelDateColumn = Nothing
								End If

								If Len(sErrorDescription) = 0 Then
									' Get the find records.
									Dim cmdFindRecords = CreateObject("ADODB.Command")
									cmdFindRecords.CommandText = "sp_ASRIntGetFindRecords3"
									cmdFindRecords.CommandType = 4 ' Stored Procedure
									cmdFindRecords.ActiveConnection = Session("databaseConnection")
									cmdFindRecords.CommandTimeout = 180

									Dim prmError = cmdFindRecords.CreateParameter("error", 11, 2) ' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmError)

									Dim prmSomeSelectable = cmdFindRecords.CreateParameter("someSelectable", 11, 2) ' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmSomeSelectable)

									Dim prmSomeNotSelectable = cmdFindRecords.CreateParameter("someNotSelectable", 11, 2) ' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmSomeNotSelectable)

									Dim prmRealSource = cmdFindRecords.CreateParameter("realSource", 200, 2, 255)	' 200=varchar, 2=output, 255=size
									cmdFindRecords.Parameters.Append(prmRealSource)

									Dim prmInsertGranted = cmdFindRecords.CreateParameter("insertGranted", 11, 2)	' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmInsertGranted)

									Dim prmDeleteGranted = cmdFindRecords.CreateParameter("deleteGranted", 11, 2)	' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmDeleteGranted)

									Dim prmTableID = cmdFindRecords.CreateParameter("tableID", 3, 1)
									cmdFindRecords.Parameters.Append(prmTableID)
									prmTableID.value = CleanNumeric(Session("tableID"))

									prmViewID = cmdFindRecords.CreateParameter("viewID", 3, 1)
									cmdFindRecords.Parameters.Append(prmViewID)
									prmViewID.value = CleanNumeric(Session("viewID"))

									Dim prmOrderID = cmdFindRecords.CreateParameter("orderID", 3, 1)
									cmdFindRecords.Parameters.Append(prmOrderID)
									prmOrderID.value = CleanNumeric(Session("orderID"))

									Dim prmParentTableID = cmdFindRecords.CreateParameter("parentTableID", 3, 1)
									cmdFindRecords.Parameters.Append(prmParentTableID)
									prmParentTableID.value = CleanNumeric(Session("parentTableID"))

									Dim prmParentRecordID = cmdFindRecords.CreateParameter("parentRecordID", 3, 1)
									cmdFindRecords.Parameters.Append(prmParentRecordID)
									prmParentRecordID.value = CleanNumeric(Session("parentRecordID"))

									Dim prmFilterDef = cmdFindRecords.CreateParameter("filterDef", 200, 1, 2147483646)
									cmdFindRecords.Parameters.Append(prmFilterDef)
									prmFilterDef.value = Session("filterDef")

									Dim prmReqRecs = cmdFindRecords.CreateParameter("reqRecs", 3, 1)
									cmdFindRecords.Parameters.Append(prmReqRecs)
									prmReqRecs.value = CleanNumeric(Session("FindRecords"))

									Dim prmIsFirstPage = cmdFindRecords.CreateParameter("isFirstPage", 11, 2) ' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmIsFirstPage)

									Dim prmIsLastPage = cmdFindRecords.CreateParameter("isLastPage", 11, 2)	' 11=bit, 2=output
									cmdFindRecords.Parameters.Append(prmIsLastPage)

									Dim prmLocateValue = cmdFindRecords.CreateParameter("locateValue", 200, 1, 2147483646)
									cmdFindRecords.Parameters.Append(prmLocateValue)
									prmLocateValue.value = Session("locateValue")

									Dim prmColumnType = cmdFindRecords.CreateParameter("columnType", 3, 2) ' 3=integer, 2=output
									cmdFindRecords.Parameters.Append(prmColumnType)

									Dim prmColumnSize = cmdFindRecords.CreateParameter("columnSize", 3, 2) ' 3=integer, 2=output
									cmdFindRecords.Parameters.Append(prmColumnSize)

									Dim prmColumnDecimals = cmdFindRecords.CreateParameter("columnDecimals", 3, 2) ' 3=integer, 2=output
									cmdFindRecords.Parameters.Append(prmColumnDecimals)

									Dim prmAction = cmdFindRecords.CreateParameter("action", 200, 1, 255)
									cmdFindRecords.Parameters.Append(prmAction)
									prmAction.value = Session("action")

									Dim prmTotalRecCount = cmdFindRecords.CreateParameter("totalRecCount", 3, 2) ' 3=integer, 2=output
									cmdFindRecords.Parameters.Append(prmTotalRecCount)

									Dim prmFirstRecPos = cmdFindRecords.CreateParameter("firstRecPos", 3, 3) ' 3=integer, 3=input/output
									cmdFindRecords.Parameters.Append(prmFirstRecPos)
									prmFirstRecPos.value = CleanNumeric(Session("firstRecPos"))

									Dim prmCurrentRecCount = cmdFindRecords.CreateParameter("currentRecCount", 3, 1)	' 3=integer, 1=input
									cmdFindRecords.Parameters.Append(prmCurrentRecCount)
									prmCurrentRecCount.value = CleanNumeric(Session("currentRecCount"))

									Dim prmColumnThousand = cmdFindRecords.CreateParameter("Use1000Separator", 11, 2) ' 3=integer, 2=output
									cmdFindRecords.Parameters.Append(prmColumnThousand)

									Dim prmDecSeparator = cmdFindRecords.CreateParameter("decSeparator", 200, 1, 255) ' 200=varchar, 1=input, 255=size
									cmdFindRecords.Parameters.Append(prmDecSeparator)
									prmDecSeparator.value = Session("LocaleDecimalSeparator")

									Dim prmDateFormat = cmdFindRecords.CreateParameter("dateFormat", 200, 1, 255) ' 200=varchar, 1=input, 255=size
									cmdFindRecords.Parameters.Append(prmDateFormat)
									prmDateFormat.value = Session("LocaleDateFormat")

									Dim prmColumnBlankIfZero = cmdFindRecords.CreateParameter("BlankIfZero", 11, 2) ' 3=integer, 2=output
									cmdFindRecords.Parameters.Append(prmColumnBlankIfZero)

									Err.Clear()
									Dim rstFindRecords = cmdFindRecords.Execute

									If (Err.Number <> 0) Then
										sErrorDescription = "The find records could not be retrieved." & vbCrLf & formatError(Err.Description)
									End If

									If Len(sErrorDescription) = 0 Then
										' Instantiate and initialise the grid. 
										Response.Write("<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridFindRecords name=ssOleDBGridFindRecords codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Rows"" VALUE=""10"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Col.Count"" VALUE=""" & rstFindRecords.fields.count & """>" & vbCrLf)
										Response.Write("	<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""MultiLine"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
										Response.Write("	<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""MaxSelectedRows"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
										Response.Write("	<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
										Response.Write("	<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Columns.Count"" VALUE=""" & rstFindRecords.fields.count & """>" & vbCrLf)

										For iLoop = 0 To (rstFindRecords.fields.count - 1)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""5600"">" & vbCrLf)
	
											If rstFindRecords.fields(iLoop).name = "ID" Then
												Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""0"">" & vbCrLf)
											Else
												Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""-1"">" & vbCrLf)
											End If
	
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Caption"" VALUE=""" & Replace(rstFindRecords.fields(iLoop).name, "_", " ") & """>" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Name"" VALUE=""" & rstFindRecords.fields(iLoop).name & """>" & vbCrLf)
				
											If (rstFindRecords.fields(iLoop).type = 131) Or (rstFindRecords.fields(iLoop).type = 3) Then
												Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""1"">" & vbCrLf)
											Else
												Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""0"">" & vbCrLf)
											End If
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Bound"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").DataField"" VALUE=""Column " & iLoop & """>" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").DataType"" VALUE=""8"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Level"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").NumberFormat"" VALUE="""">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Case"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Locked"" VALUE=""0"">" & vbCrLf)
				
											If rstFindRecords.fields(iLoop).type = 11 Then
												' Find column is a logic column.
												Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""2"">" & vbCrLf)
											Else
												Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""0"">" & vbCrLf)
											End If

											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").RowCount"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").ColCount"" VALUE=""1"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").ForeColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").BackColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").StyleSet"" VALUE="""">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Nullable"" VALUE=""1"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").Mask"" VALUE="""">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").ClipMode"" VALUE=""0"">" & vbCrLf)
											Response.Write("	<PARAM NAME=""Columns(" & iLoop & ").PromptChar"" VALUE=""95"">" & vbCrLf)
										Next

										Response.Write("	<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
										Response.Write("	<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
										Response.Write("	<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

										' JPD20020903 Fault 2316
										Response.Write("	<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
										Response.Write("</OBJECT>" & vbCrLf)

										Dim lngRowCount = 0

										' JPD 20020408 Fault 3721
										If rstFindRecords.fields.count > 0 Then
											Do While Not rstFindRecords.EOF
												' JPD20020903 Fault 2316
												Dim sAddString = ""
		
												For iLoop = 0 To (rstFindRecords.fields.count - 1)
					
													If rstFindRecords.fields(iLoop).type = 135 Then
														' Field is a date so format as such.
														' JPD20020903 Fault 2316
														'Response.Write "	<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE=""" & convertSQLDateToLocale(rstFindRecords.Fields(iLoop).Value) & """>" & vbcrlf
														sAddString = sAddString & convertSQLDateToLocale(rstFindRecords.Fields(iLoop).Value) & "	"
													ElseIf rstFindRecords.fields(iLoop).type = 131 Then
														' Field is a numeric so format as such.
														If IsDBNull(rstFindRecords.Fields(iLoop).Value) Then
															' JPD20020903 Fault 2316
															'Response.Write "	<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE="""">" & vbcrlf
															sAddString = sAddString & "	"
														Else															
															If Mid(sThousandColumns, iLoop + 1, 1) = "1" Then
																sTemp = ""
																sTemp = FormatNumber(rstFindRecords.Fields(iLoop).Value, rstFindRecords.Fields(iLoop).numericScale, True, False, True)
																sTemp = Replace(sTemp, ".", "x")
																sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
																sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
																' sAddString = sAddString & sTemp & "	"
															Else
																sTemp = ""
																sTemp = FormatNumber(rstFindRecords.Fields(iLoop).Value, rstFindRecords.Fields(iLoop).numericScale, True, False, False)
																sTemp = Replace(sTemp, ".", "x")
																sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
																sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
																' sAddString = sAddString & sTemp & "	"
															End If
								
															' NPG20090210 Fault 13249
															If Mid(sBlankIfZeroColumns, iLoop + 1, 1) = "1" And rstFindRecords.Fields(iLoop).Value = "0" Then
																sTemp = ""
															End If
								
															sAddString = sAddString & sTemp & "	"
								
														End If
													Else
														' JPD20020903 Fault 2316
														'Response.Write "	<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE=""" & rstFindRecords.Fields(iLoop).Value & """>" & vbcrlf
														If IsDBNull(rstFindRecords.Fields(iLoop).Value) Then
															sAddString = sAddString & "	"
														Else
															sAddString = sAddString & Replace(Left(rstFindRecords.Fields(iLoop).Value, 255), """", "&quot;") & "	"
														End If
													End If
												Next

												' JPD20020903 Fault 2316
												Response.Write("<INPUT type='hidden' id=txtAddString_" & lngRowCount & " name=txtAddString_" & lngRowCount & " value=""" & sAddString & """>" & vbCrLf)

												lngRowCount = lngRowCount + 1
												rstFindRecords.MoveNext()
											Loop
										End If
			
										' JPD20020903 Fault 2316
										'Response.Write "	<PARAM NAME=""Row.Count"" VALUE=""" & lngRowCount & """>" & vbcrlf
										'Response.Write "</OBJECT>" & vbcrlf

										' Release the ADO recorddim object.
										rstFindRecords.close()
										rstFindRecords = Nothing

										' NB. IMPORTANT ADO NOTE.
										' When calling a stored procedure which returns a recorddim AND has output parameters
										' you need to close the recorddim and dim it to nothing before using the output parameters. 
										If cmdFindRecords.Parameters("error").Value <> 0 Then
											sErrorDescription = "Error reading order definition."
										Else
											If cmdFindRecords.Parameters("someSelectable").Value = 0 Then
												sErrorDescription = "You do not have permission to read any of the selected order's find columns."
											End If
										End If
			
										Response.Write("<INPUT type='hidden' id=txtInsertGranted name=txtInsertGranted value=" & cmdFindRecords.Parameters("insertGranted").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtDeleteGranted name=txtDeleteGranted value=" & cmdFindRecords.Parameters("deleteGranted").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtIsFirstPage name=txtIsFirstPage value=" & cmdFindRecords.Parameters("isFirstPage").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtIsLastPage name=txtIsLastPage value=" & cmdFindRecords.Parameters("isLastPage").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtFirstColumnType name=txtFirstColumnType value=" & cmdFindRecords.Parameters("columnType").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtRecordCount name=txtRecordCount value=" & lngRowCount & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtTotalRecordCount name=txtTotalRecordCount value=" & cmdFindRecords.Parameters("totalRecCount").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtFindRecords name=txtFindRecords value=" & Session("FindRecords") & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtFirstRecPos name=txtFirstRecPos value=" & cmdFindRecords.Parameters("firstRecPos").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtCurrentRecCount name=txtCurrentRecCount value=" & lngRowCount & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtFirstColumnSize name=txtFirstColumnSize value=" & cmdFindRecords.Parameters("columnSize").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtFirstColumnDecimals name=txtFirstColumnDecimals value=" & cmdFindRecords.Parameters("columnDecimals").Value & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtCancelDateColumn name=txtCancelDateColumn value=" & fCancelDateColumn & ">" & vbCrLf)
										Response.Write("<INPUT type='hidden' id=txtGotoAction name=txtGotoAction value=" & Session("action") & ">" & vbCrLf)
			
										Session("realSource") = cmdFindRecords.Parameters("realSource").Value
									End If
	
									' Release the ADO command object.
									cmdFindRecords = Nothing
								End If
							%>
						</td>
						<td>
						</td>
					</tr>
					<%

						If Len(sErrorDescription) = 0 Then
							'
							' Get the summary fields (if required).
							'
							If Session("parentTableID") > 0 Then
								Dim cmdSummaryFields = CreateObject("ADODB.Command")
								cmdSummaryFields.CommandText = "sp_ASRIntGetSummaryFields"
								cmdSummaryFields.CommandType = 4	' Stored Procedure
								cmdSummaryFields.ActiveConnection = Session("databaseConnection")

								Dim prmHistoryTableID = cmdSummaryFields.CreateParameter("historyTableID", 3, 1)	'Type 3 = integer, Direction 1 = Input
								cmdSummaryFields.Parameters.Append(prmHistoryTableID)
								prmHistoryTableID.value = CleanNumeric(Session("tableID"))

								Dim prmParentTableID = cmdSummaryFields.CreateParameter("parentTableID", 3, 1) 'Type 3 = integer, Direction 1 = Input
								cmdSummaryFields.Parameters.Append(prmParentTableID)
								prmParentTableID.value = CleanNumeric(Session("parentTableID"))

								Dim prmParentRecordID = cmdSummaryFields.CreateParameter("parentRecordID", 3, 1)	'Type 3 = integer, Direction 1 = Input
								cmdSummaryFields.Parameters.Append(prmParentRecordID)
								prmParentRecordID.value = CleanNumeric(Session("parentRecordID"))

								Dim prmCanSelect = cmdSummaryFields.CreateParameter("canSelect", 11, 2)	'Type 11 = bit, Direction 2 = Output
								cmdSummaryFields.Parameters.Append(prmCanSelect)
	
								Err.Clear()
								Dim rstSummaryFields = cmdSummaryFields.Execute

								If (Err.Number <> 0) Then
									sErrorDescription = "The summary field definition could not be retrieved." & vbCrLf & formatError(Err.Description)
								End If

								Dim sThousSepSummaryFields As String
								Dim aSummaryFields(0, 0) As String
								Dim iTotalCount As Integer
								
								If Len(sErrorDescription) = 0 Then
									sThousSepSummaryFields = ","
									' Read the summary field definitions into an array.
									' We do this as we may be doing a lot of jumping around
									' the definitions and its easy to jump around an array than
									' a recordset.
									ReDim aSummaryFields(9, 0)
									Do While Not rstSummaryFields.EOF
										iTotalCount = UBound(aSummaryFields, 2) + 1
										ReDim Preserve aSummaryFields(9, iTotalCount)

										aSummaryFields(1, iTotalCount) = rstSummaryFields.Fields(1).Value
										aSummaryFields(2, iTotalCount) = rstSummaryFields.Fields(2).Value
										aSummaryFields(3, iTotalCount) = rstSummaryFields.Fields(3).Value
										aSummaryFields(4, iTotalCount) = rstSummaryFields.Fields(4).Value
										aSummaryFields(5, iTotalCount) = rstSummaryFields.Fields(5).Value
										aSummaryFields(6, iTotalCount) = rstSummaryFields.Fields(6).Value
										aSummaryFields(7, iTotalCount) = rstSummaryFields.Fields(7).Value
										aSummaryFields(8, iTotalCount) = rstSummaryFields.Fields(8).Value
										aSummaryFields(9, iTotalCount) = rstSummaryFields.Fields(9).Value
	
										If rstSummaryFields.Fields(9).Value Then
											sThousSepSummaryFields = sThousSepSummaryFields & CStr(rstSummaryFields.Fields(3).Value) & ","
										End If
					
										rstSummaryFields.MoveNext()
									Loop

									' Release the ADO recorddim object.
									rstSummaryFields.close()
									rstSummaryFields = Nothing

									Dim iRowCount = CLng((iTotalCount + 1) / 2)

									If iTotalCount > 0 Then
										Response.Write("				<TR height=10>" & vbCrLf)
										Response.Write("				  <TD colspan=5 height=10></TD>" & vbCrLf)
										Response.Write("				</TR>" & vbCrLf)
										Response.Write("				<TR height=10>" & vbCrLf)
										Response.Write("				  <TD colspan=5 align=center height=10>" & vbCrLf)
										Response.Write("    				<STRONG>History Summary</STRONG>" & vbCrLf)
										Response.Write("  				</TD>" & vbCrLf)
										Response.Write("				</TR>" & vbCrLf)
										Response.Write("				<TR height=10>" & vbCrLf)
										Response.Write("				  <TD colspan=5 height=10></TD>" & vbCrLf)
										Response.Write("				</TR>" & vbCrLf)

										Response.Write("				<TR height=10>" & vbCrLf)
										Response.Write("  				<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)

										Response.Write("  				<TD width=""48%"" height=10>" & vbCrLf)
										Response.Write("      			<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
									End If

									For iLoop = 1 To iRowCount
										Response.Write("   						<TR>" & vbCrLf)
										Response.Write("   							<TD nowrap=true>" & Replace(aSummaryFields(2, iLoop), "_", " ") & " :</TD>" & vbCrLf)
										Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
										Response.Write("								<TD width=""100%"">" & vbCrLf)

										If aSummaryFields(7, iLoop) = 1 Then
											' The summary control is a checkbox.
					%>
					<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
						name="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
						disabled="disabled">
					<%
					Else
						' The summary control is not a checkbox. Use a textbox for everything else.
					%>
						<input type="text" id="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
						       name="ctlSummary_<%=aSummaryFields(3, iLoop)%>_<%=aSummaryFields(4, iLoop)%>"
						       class="text textdisabled" disabled="disabled" 
							<%						If aSummaryFields(8, iLoop) = 1 Then%>
								style="width: 100%;text-align: right" />
						<% ElseIf aSummaryFields(8, iLoop) = 2 Then %> 
							style="width: 100%;text-align: center" />
						<% End If%>
						
					<%
					End If
					Response.Write("								</TD>" & vbCrLf)
					Response.Write("							</TR>" & vbCrLf)
				Next
		
				If iTotalCount > 0 Then
					Response.Write("      			</TABLE>" & vbCrLf)
					Response.Write("      		</TD>" & vbCrLf)

					Response.Write("  				<TD width=100 height=10>&nbsp;&nbsp;&nbsp;&nbsp;</TD>" & vbCrLf)

					' Do the second column now.
					Response.Write("  				<TD width=""48%"" height=10>" & vbCrLf)
					Response.Write("      			<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
				End If
				
				Dim iColumn2Index As Integer
				
				For iLoop = 1 To iRowCount
					iColumn2Index = iLoop + iRowCount
						
					If iColumn2Index <= iTotalCount Then
						Response.Write("   						<TR>" & vbCrLf)
						Response.Write("								<TD nowrap=true>" & Replace(aSummaryFields(2, iColumn2Index), "_", " ") & " :</TD>" & vbCrLf)
						Response.Write("								<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
						Response.Write("								<TD width=""100%"">" & vbCrLf)

						If aSummaryFields(7, iColumn2Index) = 1 Then
							' The summary control is a checkbox.
					%>
					<input type="checkbox" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
						name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
						disabled="disabled">
					<%
					Else
						' The summary control is not a checkbox. Use a textbox for everything else.
					%>
					<input type="text" id="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"
						name="ctlSummary_<%=aSummaryFields(3, iColumn2Index)%>_<%=aSummaryFields(4, iColumn2Index)%>"						
						<%if aSummaryFields(8, iColumn2Index) = 1 then%>
							style="width: 100%" disabled="disabled" class="text textdisabled" style="text-align: right " />
						<% elseif aSummaryFields(8, iColumn2Index) = 2 then %> 
							style="width: 100%" disabled="disabled" class="text textdisabled" style="text-align: center " />
						<%end if %>
					<%	
					End If
				End If

				Response.Write("								</TD>" & vbCrLf)
				Response.Write("							</TR>" & vbCrLf)
			Next

			If iTotalCount > 0 Then
				Response.Write("      			</TABLE>" & vbCrLf)
				Response.Write("      		</TD>" & vbCrLf)
				Response.Write("  				<TD width=20>&nbsp;&nbsp;</TD>" & vbCrLf)
					
				Response.Write("				</TR>" & vbCrLf)
			End If
		End If
			
		' NB. IMPORTANT ADO NOTE.
		' When calling a stored procedure which returns a recorddim AND has output parameters
		' you need to close the recorddim and dim it to nothing before using the output parameters. 
		Dim fCanSelect = cmdSummaryFields.Parameters("canSelect").Value

		' Release the ADO command object.
		cmdSummaryFields = Nothing

		If fCanSelect Then
			Dim cmdSummaryValues = CreateObject("ADODB.Command")
			cmdSummaryValues.CommandText = "spASRIntGetSummaryValues"
			cmdSummaryValues.CommandType = 4	' Stored Procedure
			cmdSummaryValues.ActiveConnection = Session("databaseConnection")

			Dim prmHistoryTableID2 = cmdSummaryValues.CreateParameter("historyTableID", 3, 1)	'Type 3 = integer, Direction 1 = Input
			cmdSummaryValues.Parameters.Append(prmHistoryTableID2)
			prmHistoryTableID2.value = CleanNumeric(Session("tableID"))

			Dim prmParentTableID2 = cmdSummaryValues.CreateParameter("parentTableID", 3, 1) 'Type 3 = integer, Direction 1 = Input
			cmdSummaryValues.Parameters.Append(prmParentTableID2)
			prmParentTableID2.value = CleanNumeric(Session("parentTableID"))

			Dim prmParentRecordID2 = cmdSummaryValues.CreateParameter("parentRecordID", 3, 1)	'Type 3 = integer, Direction 1 = Input
			cmdSummaryValues.Parameters.Append(prmParentRecordID2)
			prmParentRecordID2.value = CleanNumeric(Session("parentRecordID"))

			Err.Clear()
			Dim rstSummaryValues = cmdSummaryValues.Execute

			If (Err.Number <> 0) Then
				sErrorDescription = "The summary field values could not be retrieved." & vbCrLf & formatError(Err.Description)
			End If
			Dim sTempValue As String
					
			If Len(sErrorDescription) = 0 Then
				If Not (rstSummaryValues.EOF And rstSummaryValues.bof) Then
					For iLoop = 0 To (rstSummaryValues.fields.count - 1)
						If rstSummaryValues.fields(iLoop).type = 131 Then
							sTemp = "," & rstSummaryValues.fields(iLoop).name & ","

							If IsDBNull(rstSummaryValues.fields(iLoop).value) Then
								sTempValue = "0"
							Else
								sTempValue = rstSummaryValues.fields(iLoop).value
							End If

							If InStr(sThousSepSummaryFields, sTemp) > 0 Then
								sTemp = ""
								sTemp = FormatNumber(sTempValue, rstSummaryValues.Fields(iLoop).numericScale, True, False, True)
							Else
								sTemp = ""
								sTemp = FormatNumber(sTempValue, rstSummaryValues.Fields(iLoop).numericScale, True, False, False)
							End If
							sTemp = Replace(sTemp, ".", "x")
							sTemp = Replace(sTemp, ",", Session("LocaleThousandSeparator"))
							sTemp = Replace(sTemp, "x", Session("LocaleDecimalSeparator"))
							
							Response.Write("			<INPUT type='hidden' id=txtSummaryData_" & rstSummaryValues.fields(iLoop).name & " name=txtSummaryData_" & rstSummaryValues.fields(iLoop).name & " value=""" & sTemp & """>" & vbCrLf)
						Else
							Response.Write("			<INPUT type='hidden' id=txtSummaryData_" & rstSummaryValues.fields(iLoop).name & " name=txtSummaryData_" & rstSummaryValues.fields(iLoop).name & " value=""" & rstSummaryValues.fields(iLoop).value & """>" & vbCrLf)
						End If
					Next
				End If

				rstSummaryValues.close()
			End If

			rstSummaryValues = Nothing
			cmdSummaryValues = Nothing
		End If
	End If
End If
	
If Len(sErrorDescription) = 0 Then
	Response.Write("				<INPUT type='hidden' id=txtCurrentTableID name=txtCurrentTableID value=" & Session("tableID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtCurrentViewID name=txtCurrentViewID value=" & Session("viewID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtCurrentScreenID name=txtCurrentScreenID value=" & Session("screenID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & Session("orderID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtCurrentRecordID name=txtCurrentRecordID value=" & Session("recordID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtCurrentParentTableID name=txtCurrentParentTableID value=" & Session("parentTableID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtCurrentParentRecordID name=txtCurrentParentRecordID value=" & Session("parentRecordID") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtRealSource name=txtRealSource value=" & Session("realSource") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtLineage name=txtLineage value=" & Session("lineage") & ">" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtFilterDef name=txtFilterDef value=""" & Replace(Session("filterDef"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("				<INPUT type='hidden' id=txtFilterSQL name=txtFilterSQL value=""" & Replace(Session("filterSQL"), """", "&quot;") & """>" & vbCrLf)
End If

Response.Write("				<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
					%>
					<tr height="10">
						<td align="center" colspan="5" height="10">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</form>
	<form id="frmTBData" name="frmTBData">
	<%
		If CLng(Session("tableID")) = CLng(Session("TB_TBTableID")) Then
			Response.Write("				<INPUT type='hidden' id=txtTBCancelCourseDate name=txtTBCancelCourseDate value=""" & Session("lineage") & """>")
		End If
	%>
	</form>
	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>
	
	<script type="text/javascript"> find_window_onload();</script>

</div>
