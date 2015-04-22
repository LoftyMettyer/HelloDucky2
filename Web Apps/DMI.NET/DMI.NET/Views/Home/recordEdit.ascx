<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script type="text/javascript">

	//Global
	if (typeof rowWasModified === 'undefined')
		var rowWasModified = false;

	//Fault HRPRO-2953
	(function () {
		if (document.selection && document.selection.empty) {
			document.selection.empty();
		}
	})();

	$(document).on('blur', '.datepicker', function (sender) {
		if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
			OpenHR.modalMessage("Invalid date value entered");
			$(sender.target.id).focus();
		}
	});

	$(document).on('keydown', '.datepicker', function (event) {

		switch (event.keyCode) {
			case 113:    // F2 insert todays date
				$(this).datepicker("setDate", new Date())
				$(this).datepicker('widget').hide('true');
				break;
			case 37:    // LEFT --> -1 day
				//todo
				break;
			case 38:    // UPP --> -7 day
				//todo
				break;
			case 39:    // RIGHT --> +1 day
				//todo
				break;
			case 40:    // DOWN --> +7 day
				//todo
				break;
		}
	});

	function recordEdit_window_onload() {

		//public variables
		this.mavIDColumns = new Array();
		var frmRecordEditForm = OpenHR.getForm("workframe", "frmRecordEditForm");

		var fOK;
		fOK = true;
		var sErrMsg = frmRecordEditForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			// Expand the work frame and hide the option frame.
			$('#optionframe').hide();
			$('#workframe').show();
			$("#toolbarRecord").show();
			$("#toolbarRecord").click();

			$("#workframe").attr("data-framesource", "RECORDEDIT");

			var recEditCtl = document.getElementById("ctlRecordEdit"); // frmRecordEditForm.ctlRecordEdit;

			if (recEditCtl == null) {
				fOK = false;

				// The recEdit control was not loaded properly.
				OpenHR.messageBox("Record Edit control not loaded.");
				window.location = "login";
			}
		}

		if (fOK == true) {
			var frmMenuInfo = $("#frmMenuInfo")[0].children;

			var sKey = new String("photoPath_");
			sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
			var sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmRecordEditForm.txtPicturePath.value = sPath;

			sKey = new String("imagePath_");
			sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
			sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmRecordEditForm.txtImagePath.value = sPath;

			sKey = new String("olePath_");
			sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
			sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmRecordEditForm.txtOLEServerPath.value = sPath;

			sKey = new String("localolePath_");
			sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
			sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmRecordEditForm.txtOLELocalPath.value = sPath;

			//Create all tabs first...
			if (fOK == true) {
				var tabsList = $('#txtRecEditTabCaptions').val();
				if (tabsList.length > 0) {
					var aTabsList = tabsList.split('\t');
					for (var iTabCount = 0; iTabCount < aTabsList.length; iTabCount++) {
						addTabControl(iTabCount + 1);
					}
				}

			}


			if (fOK == true) {
				// Get the recEdit control to instantiate the required controls.                    
				controlCollection = frmRecordEditForm.elements;
				if (controlCollection != null) {
					txtControls = new Array();
					txtControlsCount = 0;

					//two loops here - the controlCollection was growing as controls were added, which didn't help.
					for (i = 0; i < controlCollection.length; i++) {
						sControlName = controlCollection.item(i).name;
						sControlName = sControlName.substr(0, 18);
						if (sControlName == "txtRecEditControl_") {
							//fOK = recEditCtl.addControl(controlCollection.item(i).value);
							txtControls[txtControlsCount] = controlCollection.item(i).name;
							txtControlsCount += 1;
						}

						if (fOK == false) {
							break;
						}
					}

					//Now add the form controls based on the fixed array of txtRecEditControl_ items...                        	                    
					for (i = 0; i < txtControls.length; i++) {
						txtControlValue = $("#" + txtControls[i]).val();
						var txtControlID = $("#" + txtControls[i]).attr("id");
						AddHtmlControl(txtControlValue, txtControlID, i);
					}

				}
			}

			//jQuery Functionality:
			if (fOK == true) {
				//add datepicker functionality.
				$(".datepicker").datepicker();

				//add spinner functionality
				$('.spinner').each(function () {
					var id = $(this).attr('id');
					var minvalue = $(this).attr('data-minval');
					var maxvalue = $(this).attr('data-maxval');
					var increment = $(this).attr('data-increment');
					var disabledflag = $(this).attr('data-disabled');

					$('#' + id).spinner({
						min: minvalue,
						max: maxvalue,
						step: increment,
						disabled: disabledflag,
						spin: function (event, ui) { enableSaveButton(); }
					}).on('input', function () {
						var val = this.value,
							$this = $(this),
							max = $this.spinner('option', 'max'),
							min = $this.spinner('option', 'min');
						//if (!val.match(/^\d+$/)) val = 0; //we want only number, no alpha			                
						this.value = val > max ? max : val < min ? min : val;
					}).blur(function () {
						if (isNaN(this.value)) this.value = 0;
					});

				});

				$(".spinner").spinner();

				//Loop over the "number" fields
				$(".number").each(function () {
					var control = $(this);
					control.autoNumeric('init'); //Attach autoNumeric plugin to each instance of a numeric field; this provides functionality such as masking, validate numbers, etc.
					$(control).blur(function () { //On blur, set the field to the value of the data-blankIfZeroValue attribute, set in recordEdit.js
						if ($(this).val() == 0) {
							$(this).val($(this).attr('data-blankIfZeroValue'));
						}
					});
					$(control).on("keyup", function () {
						$("#ctlRecordEdit #changed").val("false");
						enableSaveButton();
					});
				});
			}


			if (fOK == true) {
				// Set the column control values in the recEdit control.
				var sControlName;
				controlCollection = frmRecordEditForm.elements;
				if (controlCollection != null) {
					var txtControls = new Array();
					var txtControlsCount = 0;

					for (i = 0; i < controlCollection.length; i++) {
						sControlName = controlCollection.item(i).name;
						if (sControlName) {
							sControlName = sControlName.substr(0, 24);
							if (sControlName == "txtRecEditControlValues_") {
								//fOK = recEditCtl.addControlValues(controlCollection.item(i).value);
								txtControls[txtControlsCount] = controlCollection.item(i).name;
								txtControlsCount += 1;
							}
						}
						if (fOK == false) {
							break;
						}
					}

					//Now add the form control values based on the fixed array of txtRecEditControl_ items...
					for (var i = 0; i < txtControls.length; i++) {
						var txtControlValue = $("#" + txtControls[i]).val();
						addHTMLControlValues(txtControlValue);
					}
				}
			}

			if (fOK == true) {
				//.formatScreen is redundant. the 'addHtmlControl' js function replaces it and amalgamates with addControl.
				// Get the recEdit control to format itself.
				//recEditCtl.formatscreen();

				//JPD 20021021 - Added picture functionality.
				//TODO: NPG
				if (frmRecordEditForm.txtImagePath.value.length > 0) {
					var controlCollection = frmRecordEditForm.elements;
					if (controlCollection != null) {
						for (i = 0; i < controlCollection.length; i++) {
							sControlName = controlCollection.item(i).name;
							if (sControlName) {
								sControlName = sControlName.substr(0, 18);
								if (sControlName == "txtRecEditPicture_") {
									sControlName = controlCollection.item(i).name;
									//var iPictureID = new Number(sControlName.substr(18)); 
									// recEditCtl.updatePicture(iPictureID, frmRecordEditForm.txtImagePath.value + "/" + controlCollection.item(i).value);
								}
							}
						}
					}
				}
			}

			if (fOK == true) {


				// Get the data.asp to get the required data.
				//var action = document.getElementById("txtAction");
				var dataForm = OpenHR.getForm("dataframe", "frmGetData");

				if ((frmRecordEditForm.txtAction.value == "NEW" || frmRecordEditForm.txtAction.value == "COPY") &&
						(frmRecordEditForm.txtRecEditInsertGranted.value == "True")) {

					dataForm.txtAction.value = frmRecordEditForm.txtAction.value;
					dataForm.txtOriginalRecordID.value = frmRecordEditForm.txtCurrentRecordID.value;
				} else {
					dataForm.txtAction.value = "LOAD";
				}

				if (frmRecordEditForm.txtCurrentOrderID.value != frmRecordEditForm.txtRecEditOrderID.value) {
					frmRecordEditForm.txtCurrentOrderID.value = frmRecordEditForm.txtRecEditOrderID.value;
				}

				dataForm.txtCurrentTableID.value = frmRecordEditForm.txtCurrentTableID.value;
				dataForm.txtCurrentScreenID.value = frmRecordEditForm.txtCurrentScreenID.value;
				dataForm.txtCurrentViewID.value = frmRecordEditForm.txtCurrentViewID.value;
				dataForm.txtSelectSQL.value = frmRecordEditForm.txtRecEditSelectSQL.value;
				dataForm.txtFromDef.value = frmRecordEditForm.txtRecEditFromDef.value;
				dataForm.txtFilterSQL.value = "";
				dataForm.txtFilterDef.value = "";
				dataForm.txtRealSource.value = frmRecordEditForm.txtRecEditRealSource.value;
				dataForm.txtRecordID.value = frmRecordEditForm.txtCurrentRecordID.value;
				dataForm.txtOriginalRecordID.value = frmRecordEditForm.txtCurrentRecordID.value;
				dataForm.txtParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
				dataForm.txtParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
				dataForm.txtDefaultCalcCols.value = CalculatedDefaultColumns();
				OpenHR.submitForm(dataForm);
			}

			if (fOK != true) {
				// The recEdit control was not initialised properly.
				OpenHR.messageBox("Record Edit control not initialised properly.");
				window.location = "login";
			}
		}

		try {
			//  //NPG - recedit not resizing. Do it manually.
			//  var newHeight = ((frmRecordEditForm.txtRecEditHeight.value / 15));
			//  var newWidth = frmRecordEditForm.txtRecEditWidth.value / 15;
			//  // NHRD TFS 719 this division by 2 seems to work nice for getting the border around the screen nicely for screens with at least 4 rows of tabs.
			//  $("#ctlRecordEdit").height(newHeight + (document.getElementById("tabHeaders").offsetHeight / 2) + "px");
			//  $("#ctlRecordEdit").width(newWidth + "px");

			//// Mayank's code
			//  var newHeight = ((frmRecordEditForm.txtRecEditHeight.value / 15) + 20);
			//  var newWidth = frmRecordEditForm.txtRecEditWidth.value / 15;

			//  $("#ctlRecordEdit").height(newHeight + document.getElementById("tabHeaders").offsetHeight + "px");
			//  $("#ctlRecordEdit").width(newWidth + "px");


			//use zoom for IE9?

			//parent.window.resizeBy(-1, -1);
			//parent.window.resizeBy(1, 1);
		} catch (e) {
		}


		//Add 'changed' event handler to monitor for data changes    
		//checkbox
		//$("input:checkbox").change(function () {enableSaveButton();});

		//date, checkbox, text lostfocus, optiongroup, 
		$('input[id^="FI_"]').on("change", function () {
			$("#ctlRecordEdit #changed").val("false");
			enableSaveButton();
		});
		$('input[id^="FI_"]').on("keypress", function () {
			$("#ctlRecordEdit #changed").val("false");
			enableSaveButton();
		});

		$('input[id^="FI_"]').on("keyup", function () { //Keyup catches more keys than keypress (for example, Backspace)
			$("#ctlRecordEdit #changed").val("false");
			enableSaveButton();
		});

		//Text area (Notes field, etc.)
		$('textarea:not([readonly])').on("keypress", function () {
			$("#ctlRecordEdit #changed").val("false");
			enableSaveButton();
		});

		$('textarea:not([readonly])').on("keyup", function () { //Keyup catches more keys than keypress (for example, Backspace)
			$("#ctlRecordEdit #changed").val("false");
			enableSaveButton();
		});

		//need char live, spinner, dropdown, textarea,
		$('input[id^="FI_"]').on("keypress", function () {
			//TODO: check this; fires change too.....
			enableSaveButton();
		});

		//Dropdown lists
		$('select[id^="FI_"]').on("change", function () {
			$("#ctlRecordEdit #changed").val("false");
			enableSaveButton();
		});

	}

	function GoBack() {
		
		var hasChanged = menu_saveChanges("", true, false);
		var linksMainParams;
		if (Number($('#txtRecEditViewID').val()) > 0) {
			linksMainParams = '';
			linksMainParams += $('#txtRecEditTableID').val();
			linksMainParams += '!';
			linksMainParams += $('#txtRecEditViewID').val();
			linksMainParams += '_';
			linksMainParams += $('#txtCurrentRecordID').val();
		} else {
			linksMainParams = null;
		}

		if (hasChanged == 6 && !rowWasModified) { // 6 = No Change
			loadPartialView("linksMain", "Home", "workframe", linksMainParams);
			return false;
		} else if (hasChanged == 0 || rowWasModified) { // 0 = Changed, allow prompted navigation.
			OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
				if (answer == 1) { // OK
					rowWasModified = false;
					window.onbeforeunload = null;
					loadPartialView("linksMain", "Home", "workframe", linksMainParams);
					return false;
				} else {
					return false;
				}
			});
		} else
			return false;

	}

	function enableSaveButton() {
		if ($("#ctlRecordEdit #changed").val() == "false") {
			$("#ctlRecordEdit #changed").val("true");
			menu_toolbarEnableItem("mnutoolSaveRecord", true);
		}
		window.onbeforeunload = warning;
	}

	function warning() {
		return "You will lose your changes if you do not save before leaving this page.\n\nWhat do you want to do?";
	}

	function addActiveXHandlers() {
		$("#ctlRecordEdit").find("[data-columntype='lookup']").click(function () {
			ctlRecordEdit_LookupClick(this);
		});

		var ua = navigator.userAgent.toLowerCase();
		if (ua.indexOf('android') > 0) {
			$("#ctlRecordEdit").find("[data-columntype='lookup']").mousedown(function (e) {			
			e.preventDefault(); //Kill the dropdown.
			});
		}

		$("#ctlRecordEdit").find("[data-controlType='1024']").click(function () {
			//Before triggering the click, make sure that the control is not disabled
			if ($(this).attr("disabled") != "disabled") {
				ctlRecordEdit_OLEClick4(this);
			}
		});

		$("#ctlRecordEdit").find("[data-controlType='8']").click(function () {
			ctlRecordEdit_OLEClick4(this);
		});

	}

	function ctlRecordEdit_dataChanged() {
		// The data in the recEdit control has changed so refresh the menu.
		// Get menu.asp to refresh the menu.
		menu_refreshMenu();
	}

	function ctlRecordEdit_ToolClickRequest(lngIndex, strTool) {
		// The data in the recEdit control has changed so refresh the menu.
		// Get menu.asp to refresh the menu.
		menu_MenuClick(strTool);
	}

	function ctlRecordEdit_LinkButtonClick(plngLinkTableID, plngLinkOrderID, plngLinkViewID, plngLinkRecordID) {
		// A link button has been pressed in the recEdit control,
		// so open the link option page.
		menu_loadLinkPage(plngLinkTableID, plngLinkOrderID, plngLinkViewID, plngLinkRecordID);
	}

	function ctlRecordEdit_LookupClick(objLookup) {
		// A lookup button has been pressed in the recEdit control,
		// so open the lookup page. 					

		var plngColumnID = $(objLookup).attr("data-columnID");
		var plngLookupColumnID = $(objLookup).attr("data-LookupColumnID");
		var psLookupValue = $(objLookup).val();
		var pfMandatory = $(objLookup).attr("data-Mandatory");
		var pLookupFilterValueID = $(objLookup).attr("data-LookupFilterValueID");
		var pstrFilterValue = $("#ctlRecordEdit").find("[data-columnID='" + pLookupFilterValueID + "']").val();
		if (pstrFilterValue == undefined) pstrFilterValue = "";

		menu_loadLookupPage(plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue);
	}

	function ctlRecordEdit_ImageClick4(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize, pbIsReadOnly) {
		// An image has been pressed in the recEdit control,
		// so open the image find page.
		var fOK;

		fOK = true;
		//if (frmRecordEditForm.ctlRecordEdit.recordID == 0) {
		if ($("#txtCurrentRecordID").val() == 0) {
			OpenHR.messageBox("Unable to edit photo fields until the record has been saved.");
			fOK = false;
		}

		if (fOK == true) {
			//TODO Client DLL stuff
			//    if (plngOLEType < 2) {
			//        fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtPicturePath.value);
			//        if (fOK == true)
			//            window.parent.frames("menuframe").loadImagePage(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize);
			//        else
			//            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to edit photo fields as the photo path is not valid.");
			//    } else {
			//        window.parent.frames("menuframe").loadImagePage(plngColumnID, psImage, plngOLEType, plngMaxEmbedSize);
			//    }
		}
	}

	function ctlRecordEdit_OLEClick4(clickObj) {

		// An OLE button has been pressed in the recEdit control,
		// so open the OLE page.	
		//plngColumnID, psFile, plngOLEType, plngMaxEmbedSize, pbIsReadOnly
		var fOK;
		var sKey = new String('');
		fOK = true;
		var plngColumnID = $(clickObj).attr('data-columnID');
		var plngOleType = Number($(clickObj).attr('data-OleType'));
		var psFile = $(clickObj).attr('data-fileName');
		var plngMaxEmbedSize = $(clickObj).attr('data-maxEmbedSize');
		var pbIsReadOnly = $(clickObj).attr('data-readOnly');
		var frmMenuInfo = $("#frmMenuInfo")[0].children;
		var isPhoto = ($(clickObj).attr('data-controlType') == '1024');

		if ($("#txtCurrentRecordID").val() == 0) {
			OpenHR.messageBox("Unable to edit OLE fields until the record has been saved.");
			fOK = false;
		}

		if (fOK == true) {
			// Server OLE
			if (plngOleType == 1) {
				OpenHR.messageBox("This functionality is not available in 'OpenHR Web'");
			}

				// Local OLE
			else if (plngOleType == 0) {
				var messageText = "Please use your file browser to view local OLE documents.";
				var sPath = document.getElementById('frmRecordEditForm').txtOLELocalPath.value;
				if (sPath.length > 0) messageText += "\n\nYour local OLE documents can be found at: \n" + sPath;

				OpenHR.messageBox(messageText, 48);

				//TODO: Assume valid for now:    fOK = window.parent.frames("menuframe").ASRIntranetFunctions.ValidateDir(frmRecordEditForm.txtOLELocalPath.value);						
				if (fOK == true) {
					//menu_loadOLEPage(plngColumnID, psFile, plngOleType, plngMaxEmbedSize, pbIsReadOnly, isPhoto);
				} else
					OpenHR.messageBox("Unable to edit local OLE fields as the OLE (Local) path is not valid.");
			}

				// Embedded OLE
			else if (plngOleType == 2) {
				sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
				menu_loadOLEPage(plngColumnID, psFile, plngOleType, plngMaxEmbedSize, pbIsReadOnly, isPhoto);
			}

				// Linked OLE
			else if (plngOleType == 3) {
				sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
				menu_loadOLEPage(plngColumnID, psFile, plngOleType, plngMaxEmbedSize, pbIsReadOnly, isPhoto);
			}
		}
	}

	function refreshData() {
		// Get the data.asp to get the required data.
		var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
		var frmRecordEditForm = OpenHR.getForm("workframe", "frmRecordEditForm");

		frmGetDataForm.txtAction.value = "LOAD";
		frmGetDataForm.txtReaction.value = "";
		frmGetDataForm.txtCurrentTableID.value = frmRecordEditForm.txtCurrentTableID.value;
		frmGetDataForm.txtCurrentScreenID.value = frmRecordEditForm.txtCurrentScreenID.value;
		frmGetDataForm.txtCurrentViewID.value = frmRecordEditForm.txtCurrentViewID.value;
		frmGetDataForm.txtSelectSQL.value = frmRecordEditForm.txtRecEditSelectSQL.value;
		frmGetDataForm.txtFromDef.value = frmRecordEditForm.txtRecEditFromDef.value;
		frmGetDataForm.txtFilterSQL.value = frmRecordEditForm.txtRecEditFilterSQL.value;
		frmGetDataForm.txtFilterDef.value = frmRecordEditForm.txtRecEditFilterDef.value;
		frmGetDataForm.txtRealSource.value = frmRecordEditForm.txtRecEditRealSource.value;
		frmGetDataForm.txtRecordID.value = OpenHR.getForm("dataframe", "frmData").txtRecordID.value;
		frmGetDataForm.txtParentTableID.value = frmRecordEditForm.txtCurrentParentTableID.value;
		frmGetDataForm.txtParentRecordID.value = frmRecordEditForm.txtCurrentParentRecordID.value;
		frmGetDataForm.txtDefaultCalcCols.value = CalculatedDefaultColumns();
		frmGetDataForm.txtInsertUpdateDef.value = "";
		frmGetDataForm.txtTimestamp.value = "";

		data_refreshData();
	}

</script>

<link href="<%: Url.LatestContent("~/Content/spectrum.css")%>" rel="stylesheet" type="text/css" />

<div <%=session("BodyTag")%>>
	<form action="" method="post" id="frmRecordEditForm" name="frmRecordEditForm">
		<div class="absolutefull">

			<%
	
				Dim objDatabaseAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim SPParameters() As SqlParameter

				Dim sErrorDescription As String = ""
	
				' Get the page title.
				Dim prmTitle = New SqlParameter("psTitle", SqlDbType.VarChar, 500) With {.Direction = ParameterDirection.Output}
				Dim prmQuickEntry = New SqlParameter("pfQuickEntry", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmScreenID = New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}
				Dim prmViewID = New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))}

				Try
					objDatabaseAccess.ExecuteSP("sp_ASRIntGetRecordEditInfo", prmTitle, prmQuickEntry, prmScreenID, prmViewID)
		
				Catch ex As Exception
					sErrorDescription = "The page title could not be created." & vbCrLf & FormatError(ex.Message)

				End Try
	
				If Len(sErrorDescription) = 0 Then
					'Response.Write(Replace(cmdRecEditWindowTitle.Parameters("title").Value, "_", " ") & " - No activeX" & vbCrLf)        
					Response.Write("<input type='hidden' id=txtQuickEntry name=txtQuickEntry value=" & prmQuickEntry.Value & ">" & vbCrLf)
				End If
	
	
			%>

			<div class="pageTitleDiv">
				<%--<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Home'>--%>
				<a href='#'
					title='Back'
					onclick="GoBack()">
					<i class='pageTitleIcon icon-circle-arrow-left'></i>
				</a>
				<span style="margin-left: 40px; margin-right: 20px" class="pageTitle" id="RecordEdit_PageTitle">
					<%
						Response.Write(Replace(prmTitle.Value.ToString(), "_", " "))
					%>
				</span>
			</div>

			<div id="ctlRecordEdit" class="ui-widget-content" style="margin: 0 auto;">
				<ul id="tabHeaders">
				</ul>
				<input type="hidden" id="changed" value="false" />
			</div>

			<%
				'Save the page title in a hidden field for use in menu.js
				Response.Write("<input type='hidden' id='txtOriginalPageTitle' name='txtOriginalPageTitle' value='" & Replace(prmTitle.Value.ToString(), "_", " ") & "'>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentTableID name=txtCurrentTableID value=" & Session("tableID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentViewID name=txtCurrentViewID value=" & Session("viewID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentScreenID name=txtCurrentScreenID value=" & Session("screenID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & Session("orderID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentRecordID name=txtCurrentRecordID value=" & Session("recordID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtOriginalRecordID name=txtOriginalRecordID value=" & Session("recordID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentParentTableID name=txtCurrentParentTableID value=" & Session("parentTableID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentParentRecordID name=txtCurrentParentRecordID value=" & Session("parentRecordID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtLineage name=txtLineage value=" & Session("lineage") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCurrentRecPos name=txtCurrentRecPos value=" & Session("parentRecordID") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtCopiedRecordID name=txtCopiedRecordID value=''>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtRecEditTimeStamp name=txtRecEditTimeStamp value=''>" & vbCrLf)
	
				If Len(sErrorDescription) = 0 Then
			
					Try
			
						SPParameters = New SqlParameter() { _
								New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}, _
								New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))}}
			
						Dim rowScreenInfo = objDatabaseAccess.GetFromSP("sp_ASRIntGetScreenDefinition", SPParameters).Rows(0)

						Response.Write("<input type='hidden' id=txtRecEditTableID name=txtRecEditTableID value=" & Session("tableID") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditViewID name=txtRecEditViewID value=" & Session("viewID") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditHeight name=txtRecEditHeight value=" & rowScreenInfo("height") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditWidth name=txtRecEditWidth value=" & rowScreenInfo("width") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditTabCount name=txtRecEditTabCount value=" & rowScreenInfo("tabCount") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditTabCaptions name=txtRecEditTabCaptions value=""" & Replace(Replace(rowScreenInfo("tabCaptions").ToString(), "&", "&&"), """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFontName name=txtRecEditFontName value=""" & Replace(rowScreenInfo("fontName").ToString(), """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFontSize name=txtRecEditFontSize value=" & rowScreenInfo("fontSize") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFontBold name=txtRecEditFontBold value=" & rowScreenInfo("fontBold") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFontItalic name=txtRecEditFontItalic value=" & rowScreenInfo("fontItalic") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFontUnderline name=txtRecEditFontUnderline value=" & rowScreenInfo("fontUnderline") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFontStrikethru name=txtRecEditFontStrikethru value=" & rowScreenInfo("fontStrikethru") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditRealSource name=txtRecEditRealSource value=""" & Replace(rowScreenInfo("realSource").ToString(), """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditInsertGranted name=txtRecEditInsertGranted value=" & rowScreenInfo("insertGranted") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditDeleteGranted name=txtRecEditDeleteGranted value=" & rowScreenInfo("deleteGranted") & ">" & vbCrLf)
			
					Catch ex As Exception
						sErrorDescription = "The screen definition could not be read." & vbCrLf & FormatError(ex.Message)
			
					End Try
			
				End If

				If Len(sErrorDescription) = 0 Then
		
					Try

						Dim prmSelectSQL = New SqlParameter("psselectSQL", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmFromDef = New SqlParameter("psFromDef", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmOrderID = New SqlParameter("piOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("orderID"))}
		
						SPParameters = New SqlParameter() { _
							New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}, _
							New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))},
							prmSelectSQL, prmFromDef, prmOrderID}
						Dim dtControls As DataTable = objDatabaseAccess.GetFromSP("sp_ASRIntGetScreenControlsString2", SPParameters)

						Dim iloop = 1
						For Each objRow As DataRow In dtControls.Rows
							Response.Write("<input type='hidden' id=txtRecEditControl_" & iloop & " name=txtRecEditControl_" & iloop & " value=""" & Replace(objRow("controlDefinition").ToString(), """", "&quot;") & """>" & vbCrLf)
							iloop += 1
						Next

						Response.Write("<input type='hidden' id=txtRecEditSelectSQL name=txtRecEditSelectSQL value=""" & Replace(Replace(prmSelectSQL.Value.ToString(), "'", "'''"), """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditFromDef name=txtRecEditFromDef value=""" & Replace(Replace(prmFromDef.Value.ToString(), "'", "'''"), """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRecEditOrderID name=txtRecEditOrderID value=" & prmOrderID.Value.ToString() & ">" & vbCrLf)

			
			
						Dim rstScreenControlValues = objDatabaseAccess.GetFromSP("sp_ASRIntGetScreenControlValuesString" _
						, New SqlParameter("plngScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))})
			
						iloop = 1
						For Each objRow As DataRow In rstScreenControlValues.Rows
							Response.Write("<input type='hidden' id='txtRecEditControlValues_" & iloop & "' name='txtRecEditControlValues_" & iloop & "' value='" & Html.Encode(objRow("valueDefinition").ToString()) & "'>" & vbCrLf)
							iloop += 1
						Next
			
						'Add two more culture-specific hidden fields: number decimal separator and thousand separator; they will be used by the autoNumeric plugin
						Response.Write("<input type='hidden' id='txtRecEditControlNumberDecimalSeparator' name='txtRecEditControlNumberDecimalSeparator' value='" & Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator & "'>" & vbCrLf)
						Response.Write("<input type='hidden' id='txtRecEditControlNumberGroupSeparator' name='txtRecEditControlNumberGroupSeparator' value='" & Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator & "'>" & vbCrLf)
				
			
					Catch ex As Exception
						sErrorDescription = "The screen control definitions could not be read." & vbCrLf & FormatError(ex.Message)

					End Try
		
				End If

				Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
				Response.Write("<input type='hidden' id=txtRecEditFilterDef name=txtRecEditFilterDef value=""" & Replace(Session("filterDef_" & Session("tableID")), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtRecEditFilterSQL name=txtRecEditFilterSQL value=""" & Replace(Session("filterSQL_" & Session("tableID")), """", "&quot;") & """>" & vbCrLf)

				Dim objUtilities As HR.Intranet.Server.Utilities = Session("UtilitiesObject")
	
				Dim sTempPath = Server.MapPath("~/pictures")
				Dim picturesArray = objUtilities.GetPictures(Session("screenID"), CStr(sTempPath))

				For iCount = 1 To UBound(picturesArray, 2)
					Response.Write("<INPUT type='hidden' id=txtRecEditPicture_" & picturesArray(1, iCount) & " name=txtRecEditPicture_" & picturesArray(1, iCount) & " value=""" & picturesArray(2, iCount) & """>" & vbCrLf)
				Next
				objUtilities = Nothing

				'sReferringPage = Request.ServerVariables("HTTP_REFERER") 
				'iIndex = inStrRev(sReferringPage, "/")
				'if iIndex > 0 then
				'	sReferringPage = left(sReferringPage, iIndex - 1)
				'	if left(sReferringPage, 5) = "http:" then
				'		sReferringPage = mid(sReferringPage, 6)
				'	end if
				'end if
				'Response.Write "<INPUT type='hidden' id=txtImagePath name=txtImagePath value=""" & sReferringPage & """>" & vbcrlf				
			%>

			<input type='hidden' id="txtPicturePath" name="txtPicturePath">
			<input type='hidden' id="txtImagePath" name="txtImagePath">
			<input type='hidden' id="txtOLEServerPath" name="txtOLEServerPath">
			<input type='hidden' id="txtOLELocalPath" name="txtOLELocalPath">
		</div>
	</form>
	
	<div class="ui-state-error ui-corner-bottom" id="sessionWarning">		
		<p style="font-size: small;">Your session will time-out in </p>
		<p id="timerText"></p>
		<p style="font-size: small;">click <a onclick="resetSession();" href="#">here</a> to renew it</p>
	</div>

</div>


<script type="text/javascript">
	recordEdit_window_onload();
	//must run after onload (which populates the screen)
	addActiveXHandlers();

	//Set up the session timeout counter. 
	var mins = <%:Session.Timeout%>;  //Set the number of minutes you need
	var secs = mins * 60;
	var currentSeconds = 0;
	var currentMinutes = 0;
	var sessionTimer = setTimeout('Decrement()', 1000);

	function Decrement(newVal) {
		if (Number(newVal) > 0) {
			secs = newVal * 60;
			$("#sessionWarning").hide();
			return false;
		}

		currentMinutes = Math.floor(secs / 60);
		currentSeconds = secs % 60;
		if (currentSeconds <= 9) currentSeconds = "0" + currentSeconds;
		secs--;
		try {
			if (secs < 300) $("#sessionWarning").show();		//show countdown for the last 5 minutes.
			document.getElementById("timerText").innerHTML = currentMinutes + ":" + currentSeconds; //Set the element id you need the time put into.
			if (secs !== -1) setTimeout('Decrement()', 1000);
		} catch (e) {
			//do nothing if this fails - we've probably navigated away and the elements no longer exist. That's the trouble with using 1 second delays.
		}

	}

	function resetSession() {
		var mins = <%:Session.Timeout%>; 
		$.post('RefreshSession', function () {});
		Decrement(mins);
	}

	$(document).ready(function () {

		//if (controls with a Negative number) {
		//	$()
		//}
		// Harry's code
		var newWidth = $("#txtRecEditWidth").val() / 14.8;
		$("#ctlRecordEdit").width(newWidth + "px");
		$("#ctlRecordEdit").css("overflow", "hidden");

		var tabheight = Number($("#tabHeaders").height());
		if (tabheight < 40) {
			tabheight = 40;
		}
		var newHeight = (Number($("#txtRecEditHeight").val()) / Number(15.2)) + tabheight - 40;
		$("#ctlRecordEdit").height(newHeight + "px");


		if (menu_isSSIMode() && (window.currentLayout != "winkit")) {
			//Only for SSI mode view, zoom in on recedit until it reaches full screen, or twice it's original size.
			var screenWidth = document.getElementById("workframeset").offsetWidth;
			var screenHeight = document.getElementById("workframeset").offsetHeight;
			var scaleFactor = 0;

			//Calculate the scale factor
			if ((screenWidth / newWidth) < (screenHeight / newHeight)) {
				//use width as factor
				scaleFactor = (screenWidth * .9) / newWidth;
			} else {
				//use height
				scaleFactor = (screenHeight * .9) / newHeight;
			}

			scaleFactor = (scaleFactor * 0.8);

			//Limit the scale factor to 1.5x
			scaleFactor = Math.min(scaleFactor, 1.5);

			//$("#ctlRecordEdit").css("-webkit-transform", "scale(" + scaleFactor + ")");
			//$("#ctlRecordEdit").css("-webkit-transform-origin", "50% top");
			//$("#ctlRecordEdit").css("-moz-transform", "scale(" + scaleFactor + ")");
			//$("#ctlRecordEdit").css("-moz-transform-origin", "50% top");
			//$("#ctlRecordEdit").css("transform", "scale(" + scaleFactor + ")");
			//$("#ctlRecordEdit").css("transform-origin", "50% top");
		}

		var toolMfRecord = getCookie('toolMFRecord');
		$('#mnutoolMFRecord').removeClass('toolbarButtonOn');
		if (toolMfRecord == 'true') toggleMandatoryColumns(true);
	});

</script>
