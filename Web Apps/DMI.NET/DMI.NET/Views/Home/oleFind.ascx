<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>
<%' OLE TYPES:  0 = Local OLE, 1 = Server OLE, 2 = Embedded OLE, 3 = Linked OLE. %>


<%Dim sDialogTitle As String
	If Session("optionIsPhoto") = "true" Then
		sDialogTitle = "Select Image File"
	Else
		sDialogTitle = "Select Document"
	End If
	
	If Session("optionOLEReadOnly") = "true" Then sDialogTitle &= " (Read Only)"
%>

<script type='text/javascript'>
	function oleFind_window_onload() {

		var fOK;
		fOK = true;
		var frmFindForm = document.getElementById('frmFindForm');
		var sErrMsg = frmFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.MessageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmFindForm.cmdCancel.focus();
			var frmMenu = OpenHR.getForm("menuframe", "frmMenuInfo");
			var frmGotoOption = document.getElementById('frmFindForm');
			var sKey = '';
			var sPath = '';

			sKey = new String("olePath_");
			sKey = sKey.concat(frmMenu.txtDatabase.value);
			sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmFindForm.txtOLEServerPath.value = sPath;

			sKey = new String("localolePath_");
			sKey = sKey.concat(frmMenu.txtDatabase.value);
			sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmFindForm.txtOLELocalPath.value = sPath;

			sKey = new String("photoPath_");
			sKey = sKey.concat(frmMenu.txtDatabase.value);
			sPath = OpenHR.GetRegistrySetting("HR Pro", "DataPaths", sKey);
			frmFindForm.txtPicturePath.value = sPath;
			
			if (frmGotoOption.txtOLEType.value == 0) {
				var messageText = "Please use your file browser to view local OLE documents.";
				if (sPath.length > 0) messageText += "\n\nYour local OLE documents can be found at: \n" + sPath;

				OpenHR.messageBox(messageText, 48);
				return false;
			}

			// Expand the option frame and hide the work frame.
			$("#optionframe").attr("data-framesource", "SELECTOLE");
			$("#optionframe").dialog({
				title: '<%=sDialogTitle%>',
				autoOpen: false,
				modal: true,
				width: 'auto',
				height: 'auto',
				close: Cancel,
				closeOnEscape: false,
				open: function (event, ui) {
					$(".ui-dialog-titlebar-close", ui.dialog || ui).hide();
					$("#ssOleDBGrid").jqGrid('setGridWidth', $("#optionframe").width() - 30);
				},
				resize: function () { //resize the grid to the height of its container.		
					$("#ssOleDBGrid").jqGrid('setGridWidth', $("#optionframe").width() - 30);
				}
			});

			if (frmGotoOption.txtOLEType.value < 2) {
				// Populate the grid with the files in the specified picture folder.
				PopulateGrid();

				if (rowCount() > 0) {
					if (frmFindForm.txtFile.value.length > 0) {
						// Try to select the current record.
						locateRecord(frmFindForm.txtFile.value, true);
					} else {
						// Select the top row.
						moveFirst();
					}
				}
			} else {
				if (frmGotoOption.txtOLEFile.value == "") {
					button_disable(frmFindForm.cmdEmbed, ((frmFindForm.txtOLEReadOnly.value == 'true') ||
						(frmFindForm.txtOLEMaxEmbedSize.value == 0)));
					button_disable(frmFindForm.cmdLink, (frmFindForm.txtOLEReadOnly.value == 'true'));
					button_disable(frmFindForm.cmdEdit, true);
					button_disable(frmFindForm.cmdProperties, true);
					button_disable(frmFindForm.cmdRemove, true);
					setASRIntOLE1_FileName("");
				} else {
					button_disable(frmFindForm.cmdEdit, false);
					button_disable(frmFindForm.cmdEmbed, true);
					button_disable(frmFindForm.cmdLink, true);
					button_disable(frmFindForm.cmdRemove, (frmFindForm.txtOLEReadOnly.value == 'true'));
					setASRIntOLE1_FileName(frmGotoOption.txtOLEJustFileName.value);
				}
			}
		
			refreshControls();
			menu_disableMenu();
		}

		return false;
	}

	function setASRIntOLE1_FileName(newFilename) {
		//replicates activeX method 'FileName'.
		var frmGotoOption = document.getElementById('frmFindForm');
		var oleCaption;
		switch (Number(frmGotoOption.txtOLEType.value)) {
			case 0:
				oleCaption = 'Local file: ' + newFilename;
				$('#oleCaption h3').text(oleCaption);
				break;
			case 1:
				oleCaption = 'Server file: ' + newFilename;
				$('#oleCaption h3').text(oleCaption);
				break;
			case 2:
				if (newFilename == '') {
					oleCaption = 'Empty';
				} else {
					oleCaption = 'Embedded file: ' + newFilename;
				}
				$('#oleCaption h3').text(oleCaption);
				break;
			case 3:
				if (newFilename == '') {
					$('#tdDescription h6').text('');
					$('#oleCaption h3').html('Empty');
				} else {

					if ("ActiveXObject" in window){
						$('#tdDescription h6').text('Right-click the link below and choose \'Save As...\' to download this file.');
						$('#oleCaption h3').html('<a title="(Right-click this link and choose \'Save As...\' to download this file.)" target="submit-iframe" href="' + $('#txtOLEFile').val() + '">Linked file: ' + newFilename + '</a>');
					} else {
						//Non-IE browsers
						$('#oleCaption h3').html('Linked file: ' + newFilename);
					}
				}
				break;
			default:
				oleCaption = '';
				$('#oleCaption h3').text(oleCaption);
				break;
		}
	}
</script>

<script type="text/javascript">

	function moveFirst() {
		$("#ssOleDBGrid").jqGrid('setSelection', 1);
	}

	function PopulateGrid() {
		var lngOleType;
		var frmFindForm = document.getElementById('frmFindForm');

		lngOleType = frmFindForm.txtFFOLEType.value;

		if (lngOleType < 2) {

			// Clear the current contents of the grid.
			$("#ssOleDBGrid").jqGrid('GridUnload');

			// Server OLE
			if (lngOleType == 1)
				//fc = new Enumerator(window.parent.frames("menuframe").ASRIntranetFunctions.FolderList(frmFindForm.txtOLEServerPath.value).Files);
				FolderList(frmFindForm.txtOLEServerPath.value);

				// Local OLE
			else if (lngOleType == 0)
				//fc = new Enumerator(window.parent.frames("menuframe").ASRIntranetFunctions.FolderList(frmFindForm.txtOLELocalPath.value).Files);
				FolderList(frmFindForm.txtOLELocalPath.value);
		}
	}

	function rowCount() {
		return $("#ssOleDBGrid tr").length - 1;
	}

	function bookmarksCount() {
		return 1;
	}


	function FolderList(pstrLocation) {
		//use AJAX to return array of files in the OLE path - server-side only...
		//this only works because we convert the returned values to json for jqGrid.
		if (pstrLocation.length > 0) {

			$.ajax({
				url: "<%: Html.Raw(Url.Action("FolderList", "Home"))%>",
				type: "POST",
				async: false,
				data: { folderPath: pstrLocation },
				dataType: "json",
				success: function(data) {
					var colNames = ["filename"];
					var colModel = [{ name: 'filename' }];
					var colData = [];

					$.each(data, function(k, v) {
						colData.push({ filename: v });
					});

					//create the column layout:
					$("#ssOleDBGrid").jqGrid({
						datatype: "local",
						data: colData,
						colNames: colNames,
						colModel: colModel,
						autowidth: true
					});

				},
				error: function () {
					//assume invalid path.
					$('#ssOleDBGrid').html('<tr><td align="center"><h3>Server OLE Path is unavailable :<br/></h3><h2>' + pstrLocation + '</h2></td></tr>');					
				},
			});
		} else {
			//No path set.
			$('#ssOleDBGrid').html('<tr><td align="center"><h3>Server OLE Path has not been set</h3></td></tr>');
		}
	}

	function Select() {		
		var frmFindForm = document.getElementById('frmFindForm');
		var frmGotoOption = document.getElementById('frmFindForm');
		if (bookmarksCount() > 0) {
			$("#optionframe").dialog("destroy");
			frmGotoOption.txtGotoOptionColumnID.value = frmFindForm.txtOptionColumnID.value;
			frmGotoOption.txtGotoOptionFile.value = selectedValue();
			frmGotoOption.txtGotoOptionAction.value = "SELECTOLE";
			frmGotoOption.txtGotoOptionPage.value = "emptyoption";
			OpenHR.submitForm(frmGotoOption);

			setTimeout(function () { // Delay for Chrome
				loadEmptyOption();
			}, 100);

		}
	}

	function Clear() {
		var frmFindForm = document.getElementById('frmFindForm');
		var frmGotoOption = document.getElementById('frmFindForm');

		$("#optionframe").dialog("destroy");

		frmGotoOption.txtGotoOptionColumnID.value = frmFindForm.txtOptionColumnID.value;
		frmGotoOption.txtGotoOptionFile.value = "";
		frmGotoOption.txtGotoOptionAction.value = "SELECTOLE";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
		
		setTimeout(function () { // Delay for Chrome
			loadEmptyOption();
		}, 100);
	}

	function Cancel() {		
		var bOK;
		var iAnswer;
		var frmFindForm = document.getElementById('frmFindForm');
		var frmGotoOption = document.getElementById('frmFindForm');

		bOK = true;

		if (frmGotoOption.txtOLEType.value > 1) {
			if (frmFindForm.cmdSelect.disabled == false) {
				iAnswer = OpenHR.messageBox("All changes will be lost. Are you sure you want to cancel?", 36);
				if (iAnswer != 6) {
					bOK = false;
				}
			}
		}

		if (bOK == true) {
			//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
			$("#optionframe").dialog("destroy");

			frmGotoOption.txtGotoOptionAction.value = "CANCEL";
			frmGotoOption.txtGotoOptionPage.value = "emptyoption";
			OpenHR.submitForm(frmGotoOption);

			setTimeout(function () { // Delay for Chrome
				loadEmptyOption();
			}, 100);

		}
	}

	//This function is called when frmFindForm is submitted...
	$('#frmFindForm').submit(function (e) {
		var fOK;
		fOK = true;
		var frmFindForm = document.getElementById('frmFindForm');
		var frmGotoOption = document.getElementById('frmFindForm');

		if (frmGotoOption.txtOLEType.value < 2) {
			e.preventDefault();
			this.submit();

			setTimeout(function () { // Delay for Chrome
				$('#ssOleDBGridRow').parent().show();
				$('#fileUpload').hide();
				$('#linkUploadCaption').hide();
				$('#cmdAdd').show();
				$('#cmdEdit').show();
				$('#cmdClear').show();
				$('#cmdSelect').show();
				$('#cmdUpload2').hide();

				//reload the file list
				PopulateGrid();
			}, 100);

		} else {

			// If not blank and ole type is embedded
			if ((frmGotoOption.txtOLEJustFileName.value != "") && (frmGotoOption.txtOLEType.value == 2) && (frmGotoOption.txtOLEIsNew.value == "True")) {
				if (frmGotoOption.txtOLECommit.value == 1) {
					try {
						frmGotoOption.txtOLEFile.value = $('#filSelectFile').val();
						frmGotoOption.txtOLEEncryption.value = true;

					} catch (e) {
						OpenHR.messageBox("Unable to save your document.\nContact your system administrator.", 16);
						fOK = false;
					}
				}
			}

			// Pass the new filename in with the text to flag it as a linked file
			if (fOK == true) {
				$("#optionframe").dialog("destroy");

				//frmGotoOption.txtGotoOptionFile.value = frmGotoOption.txtOLEJustFileName.value;
				if (frmGotoOption.txtOLEType.value == 3) {
					frmGotoOption.txtGotoOptionFile.value = frmGotoOption.txtOLEFile.value + "::LINKED_OLE_DOCUMENT::";
				} else {
					frmGotoOption.txtGotoOptionFile.value = frmGotoOption.txtOLEJustFileName.value + "::EMBEDDED_OLE_DOCUMENT::";
				}

				frmGotoOption.txtGotoOptionColumnID.value = frmFindForm.txtOptionColumnID.value;

				if (frmGotoOption.txtOLECommit.value == 1) {
					frmGotoOption.txtGotoOptionAction.value = "LINKOLE";
				} else {
					frmGotoOption.txtGotoOptionAction.value = "";
				}

				frmGotoOption.txtGotoOptionPage.value = "emptyoption";
				
				recEdit_setData($('#txtOptionColumnID').val(), $('#txtFile').val());
				//TODO: 
				recEdit_setTimeStamp('<%=session("timestamp")%>');

				e.preventDefault();
				this.submit();

				setTimeout(function () { // Delay for Chrome
					loadEmptyOption();
				}, 100);

			} else {
				button_disable(frmFindForm.cmdSelect, true);
				$('#cmdSelect').button('disable');
				return false;
			}
		}
		return false;
	});


	//This function replaces the 'Response.Redirect('emptyoption') in the controller,
	//submitting it to the correct 'optionframe' div.
	function loadEmptyOption() {
		$.ajax({
			url: 'emptyoption',
			type: "POST",
			dataType: 'html',
			async: true,
			success: function (html) {
				try {
					$('#optionframe').html('');
					$('#optionframe').html(html);
				} catch (e) { }
			}
		});
	}

	function showFileUpload(linkType) {
		$('#txtFFOLEType').val(linkType);
		$('#oleCaption').hide();
		$('#fileUpload').show();
	}

	// Embed or link
	function EmbedLink() {

		var plngOleType = $('#txtFFOLEType').val();	// ($('input[name="uploadType"]:checked').val() == "embed" ? 2 : 3);		
		var sFile;
		var bOK;
		var lngFileSize;
		var frmGotoOption = document.getElementById('frmFindForm');
		var frmFindForm = document.getElementById('frmFindForm');

		bOK = true;
		lngFileSize = 0;

		// Select a file
		var filSelectFile = document.getElementById('filSelectFile');
		//filSelectFile.click();		

		// Get the selected file name.
		sFile = new String(filSelectFile.value);

		var bIsPhoto = (document.getElementById("txtIsPhoto").value == "true");

		if (bIsPhoto) {
			//validate Photo Picture Types
			//VB6 types only :(		
			var fileExtension = OpenHR.GetFileExtension(filSelectFile.value).toLocaleLowerCase();
			var validFileExtensions = ["jpg", "bmp", "gif"];
			if (validFileExtensions.indexOf(fileExtension) == -1) {
				//invalid extension
				alert("Invalid image type.\n\nOnly .JPG, .BMP and .GIF images are accepted.");
				return false;
			}
		}

		// Check that the filename/unc isn't too long
		if (plngOleType != 2) {
			if ((sFile.length > 0) && (plngOleType > 1)) {
				var sMessage = new String(OpenHR.CheckOLEFileNameLength(filSelectFile.value));
				if (sMessage.length > 0) {
					OpenHR.messageBox(sMessage, 48);
					bOK = false;
				}
			}
		}

		if ((sFile.length > 0) && (bOK == true)) {
			// Load the submit form
			frmGotoOption.txtOLEJustFileName.value = OpenHR.getFileNameOnly(filSelectFile.value);
			frmGotoOption.txtOLEFile.value = OpenHR.ConvertToUNC(filSelectFile.value);
			frmGotoOption.txtOLEFileUNCPath.value = OpenHR.GetPathOnly(frmGotoOption.txtOLEFile.value, false);
			frmGotoOption.txtOLEType.value = plngOleType;
			frmGotoOption.txtOLEEncryption.value = false;
			frmGotoOption.txtOLECommit.value = 1;

			var datelastmodified;
			//IE9 won't do this:
			try {
				var file = document.getElementById('filSelectFile').files[0];
				datelastmodified = file.lastModifiedDate;
			} catch (e) {
				datelastmodified = Date.today();
			}

			//TODO:
			datelastmodified = '01/01/2001 00:00';

			frmGotoOption.txtOLEModifiedDate.value = datelastmodified;		// window.parent.frames("menuframe").ASRIntranetFunctions.FileLastModified(frmGotoOption.txtOLEFile.value);
			
			if (plngOleType < 2) {
				//We're done for server-side oles.
				return false;
			}


			// Update the display
			//TODO: Set these properties on record edit?
			//frmFindForm.ASRIntOLE1.DMIBackColor = 16777215;
			//frmFindForm.ASRIntOLE1.DisplayFileName = window.parent.frames("menuframe").ASRIntranetFunctions.getFileNameOnly(filSelectFile.value);
			//frmFindForm.ASRIntOLE1.FileName = filSelectFile.value;
			setASRIntOLE1_FileName(OpenHR.getFileNameOnly(filSelectFile.value));
			//frmFindForm.ASRIntOLE1.OLEType = plngOLEType;
			//frmFindForm.ASRIntOLE1.IsFileEncrypted = false;
			//frmFindForm.ASRIntOLE1.DisplayFileImage();
			button_disable(frmFindForm.cmdSelect, (frmFindForm.txtOLEReadOnly.value == 'true'));
			$('#cmdSelect').button(frmFindForm.txtOLEReadOnly.value == 'true' ? 'disable' : 'enable');
			button_disable(frmFindForm.cmdProperties, false);
			button_disable(frmFindForm.cmdRemove, (frmFindForm.txtOLEReadOnly.value == 'true'));

			//Disable download button for Linked files.
			//button_disable(frmFindForm.cmdEdit, (frmFindForm.txtOLEType.value == 3));

			//Disable download button for newly embedded files.
			button_disable(frmFindForm.cmdEdit, true);

			// Disable the embed/link buttons
			button_disable(frmFindForm.cmdEmbed, true);
			button_disable(frmFindForm.cmdLink, true);

			// Change the remove button text
			if (plngOleType == 3)
			{ //frmFindForm.cmdRemove.value = "Unlink"; 
			}
			else
			{ //frmFindForm.cmdRemove.value = "Delete"; 
			}

			//Hide the file input box
			$('#oleCaption').show();
			$('#fileUpload').hide();

		}
		return false;
	}

	// Function to display the file properties
	function Properties() {

		var sPropertiesMsg;
		var sCaption;
		var sSize;
		var frmGotoOption = document.getElementById('frmFindForm');

		sSize = frmGotoOption.txtOLEFileSize.value; // window.parent.frames("menuframe").ASRIntranetFunctions.FileSize(frmGotoOption.txtOLEFile.value);
		sPropertiesMsg = "File : " + frmGotoOption.txtOLEJustFileName.value;
		sPropertiesMsg = sPropertiesMsg + "\nSize : " + sSize;

		if (frmGotoOption.txtOLEType.value == 3) {
			sCaption = "Linked File properties";
			sPropertiesMsg = sPropertiesMsg + "\nLocation : " + frmGotoOption.txtOLEFileUNCPath.value;
		}
		else {
			sCaption = "Embedded File properties";
		}

		sPropertiesMsg = sPropertiesMsg + "\nLast Modified : " + frmGotoOption.txtOLEModifiedDate.value; // window.parent.frames("menuframe").ASRIntranetFunctions.FileLastModified(frmGotoOption.txtOLEFile.value);

		if (sPropertiesMsg.length > 0) {
			OpenHR.messageBox(sPropertiesMsg, 64, sCaption);
		}

	}
	
	// Remove the embedded document or linked file
	function Remove() {
		var iAnswer;
		var lngOleType;
		var sMessage;
		var frmGotoOption = document.getElementById('frmFindForm');
		var frmFindForm = document.getElementById('frmFindForm');

		lngOleType = frmGotoOption.txtOLEType.value;

		if (lngOleType == 3) {
			sMessage = "Are you sure you want to unlink " + frmGotoOption.txtOLEJustFileName.value + "?";
		}
		else {
			sMessage = "Are you sure you want to delete " + frmGotoOption.txtOLEJustFileName.value + "?";
		}

		iAnswer = OpenHR.messageBox(sMessage, 36);

		if (iAnswer == 6) {

			var columnID = frmFindForm.txtOptionColumnID.value;

			// Clear the values
			frmGotoOption.txtOLEJustFileName.value = "";
			frmGotoOption.txtOLEFile.value = "";
			frmGotoOption.txtOLEType.value = 3;
			frmGotoOption.txtOLEIsNew.value = "True";
			frmGotoOption.txtOLECommit.value = 1;
			if (columnID.length > 0) {
				//FI_3002_8_384

				$("#ctlRecordEdit").find("[data-columnID='" + columnID + "']").each(function () {
					// Update the display
					//All these ASRIntOLE1 values are stored on the control now.
					//frmFindForm.ASRIntOLE1.DMIBackColor = 16777215;
					//frmFindForm.ASRIntOLE1.FileName = "";
					$(this).attr('data-fileName', '');
					setASRIntOLE1_FileName('');
					//frmFindForm.ASRIntOLE1.OLEType = 3;
					$(this).attr('data-OleType', '3');
					//frmFindForm.ASRIntOLE1.DisplayFileImage();
				});

				button_disable(frmFindForm.cmdSelect, false);
				$('#cmdSelect').button('enable');
				button_disable(frmFindForm.cmdEdit, true);
				button_disable(frmFindForm.cmdProperties, true);
				button_disable(frmFindForm.cmdRemove, true);
				button_disable(frmFindForm.cmdLink, false);

				if (frmFindForm.txtOLEMaxEmbedSize.value > 0) {
					button_disable(frmFindForm.cmdEmbed, false);
				}
			}

		}

	}


	function Edit() {				
		var lngOleType;
		//var bFileEncrypted;
		//var bIsReadOnly;
		var path;
		//var sFile;
		var frmGotoOption = document.getElementById('frmFindForm');
		//var frmFindForm = document.getElementById('frmFindForm');

		lngOleType = frmGotoOption.txtOLEType.value;
		//bIsReadOnly = (frmFindForm.txtOLEReadOnly.value == 'true');

		// Server
		if (lngOleType == 1) {
			//bFileEncrypted = false;

			//Download the file!
			var dummyurl = '<%: Html.Raw(Url.Action("DownloadFile", "Home", New With {.filename = "-1", .serverpath = "-2"}))%>';
			path = dummyurl.replace("-1", selectedValue()).replace("-2", $('#frmFindForm #txtOLEServerPath').val());
			window.location.href = path;

			OpenHR.messageBox("Note: You are about to download a COPY of this document.\nIf you make changes to it, you must upload it again.\n\nClick OK to continue.", 48, "Document download");

			return false;
		}

			// Local
		else if (lngOleType == 0) {
			//should never get here now - defunct.
			return false;
		}

			// Embedded
		else if (lngOleType == 2) {
			//sFile = frmGotoOption.txtOLEJustFileName.value; // frmFindForm.ASRIntOLE1.FileName;
			//bFileEncrypted = frmGotoOption.txtOLEEncryption.value;
			frmGotoOption.txtOLECommit.value = 1;
		}

			// Linked
		else if (lngOleType == 3) {
			//sFile = frmGotoOption.txtOLEJustFileName.value; // frmFindForm.ASRIntOLE1.FileName;
			//bFileEncrypted = false;
			frmGotoOption.txtOLECommit.value = 0;
			alert("Right-click or option-click the link shown and choose 'Save As...' to download this file.");

			return false;
		}

		//window.parent.frames("menuframe").ASRIntranetFunctions.CurrentSessionKey = frmGotoOption.txtOLESession.value
		//ShowWait("System Locked...");
		//window.parent.frames("menuframe").ASRIntranetFunctions.editFile(sFile, bFileEncrypted, frmGotoOption.txtOLEJustFileName.value, bIsReadOnly);
		//extract and download the file.
		
		path = '<%: Html.Raw(Url.Action("EditFile", "Home", New With {.plngRecordID = CInt(Session("optionRecordID")), .plngColumnID = CInt(Session("optionColumnID")), .pstrRealSource = Session("realSource")}))%>';

		window.location.href = path;

		OpenHR.messageBox("Note: You are about to download a COPY of this document.\nIf you make changes to it, you must upload it again.\n\nClick OK to continue.", 48, "Document download");

		return false;

		//button_disable(frmFindForm.cmdSelect, (frmFindForm.txtOLEReadOnly.value == 'true'));

		//// If the OLE type is a link we can't control whether the changes are committed.
		//if (lngOleType == 3) {
		//	button_disable(frmFindForm.cmdCancel, (bIsReadOnly != true));
		//}

		//CloseWait();

	}


	function Add() {
		
		$('#ssOleDBGridRow').parent().hide();
		$('#fileUpload').show();
		$('#linkUploadCaption').show();
		$('#cmdAdd').hide();
		$('#cmdEdit').hide();
		$('#cmdClear').hide();
		$('#cmdSelect').hide();
		$('#cmdUpload2').show();

		return false;

		//var sFile;
		//var sFileFolder;
		//var frmFindForm = document.getElementById('frmFindForm');
		//var frmGotoOption = document.getElementById('frmFindForm');
		//var fileAddFile = document.getElementById('fileAddFile');

		//// Clear the current contents of the file object.
		//fileAddFile.value = "";

		//// Display the file selection popup.
		//fileAddFile.click();

		//// Get the selected file name.
		//sFile = new String(fileAddFile.value);
		//if (sFile.length > 0) {
		//	if (frmGotoOption.txtOLEType.value == 1) {
		//		var tmpForm = document.getElementById('AddForm');
		//		var tmpActionValue = frmGotoOption.txtGotoOptionAction.value;
		//		frmGotoOption.txtGotoOptionAction.value = "UPLOAD";
		//	}
		//}
	}

	/* Return the value of the record selected in the find form. */
	function selectedValue() {
		var sValue;

		//if (frmFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
		//	sValue = frmFindForm.ssOleDBGrid.Columns(0).Value;
		//}
		var selRowId = $('#ssOleDBGrid').jqGrid('getGridParam', 'selrow');
		sValue = $("#ssOleDBGrid").jqGrid('getCell', selRowId, 'filename');

		return (sValue);
	}

	/* Sequential search the grid for the required OLE. */

	function locateRecord(psFileName) {
		//select the grid row that contains the record with the passed in ID.
		var rowNumber = $("#ssOleDBGrid td").filter(function () {
			return $(this).text() == psFileName;
		}).closest("tr").attr("id");

		if (rowNumber >= 0) {
			$("#ssOleDBGrid").jqGrid('setSelection', rowNumber);
		} else {
			$("#ssOleDBGrid").jqGrid('setSelection', 1);
		}
	}

	function refreshControls() {
		
		var frmFindForm = document.getElementById('frmFindForm');
		var frmGotoOption = document.getElementById('frmFindForm');
		
		if (frmFindForm.txtFFOLEType.value < 2) {
			
			if (rowCount() > 0) {
				if (bookmarksCount() > 0) {
					button_disable(frmFindForm.cmdEdit, (frmFindForm.txtOLEReadOnly.value == 'true'));
					button_disable(frmFindForm.cmdSelect, (frmFindForm.txtOLEReadOnly.value == 'true'));
				}
				else {
					button_disable(frmFindForm.cmdEdit, true);
					button_disable(frmFindForm.cmdSelect, true);
				}
			}
			else {				
				button_disable(frmFindForm.cmdEdit, (frmFindForm.txtOLEReadOnly.value == 'true') );
				button_disable(frmFindForm.cmdSelect, (frmFindForm.txtOLEReadOnly.value == 'true'));				
			}
			
			//if no path set, disable all buttons except cancel.
			var serverPathMessage = $('#ssOleDBGrid h3').text();
			if ((frmFindForm.txtOLEServerPath.value.length == 0) || (serverPathMessage.length > 0)) {
				button_disable(frmFindForm.cmdAdd, true);
				button_disable(frmFindForm.cmdEdit, true);
				button_disable(frmFindForm.cmdClear, true);
				button_disable(frmFindForm.cmdSelect, true);
			}
		}
		else {			
			button_disable(frmFindForm.cmdEdit, ((frmGotoOption.txtOLEFile.value == "") || (frmFindForm.txtOLEType.value == 3)));
			$('#oleCaption').show();
			$('#fileUpload').hide();
		}


		if (frmFindForm.txtOLEReadOnly.value == 'true')
			frmFindForm.cmdEdit.value = "View";

		//Disabled Link/Unlink buttons for non-IE browsers. (FF and Chrome don't support file upload paths)
		if (!("ActiveXObject" in window)) {
			button_disable(frmFindForm.cmdLink, true);
			if(frmFindForm.cmdRemove.value == 'Unlink') button_disable(frmFindForm.cmdRemove, true);
		}

	}
</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
	<form action="oleFind_Submit" method="post" id="frmFindForm" name="frmFindForm" enctype="multipart/form-data" target="submit-iframe">

		<table class="outline aligncenter cellpadding5 cellspace0">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible cellspace0 cellpadding0">
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<% 
								If Session("optionOLEType") < 2 Then
							%>
							<td>
								<div id="ssOleDBGridRow">
									<table id="ssOleDBGrid" name="ssOleDBGrid" style="HEIGHT: 100%; LEFT: 0; TOP: 0; WIDTH: 100%"></table>
								</div>
							</td>
							<%
							Else
							%>
							<td id="oleCaption" style="display: none;text-align: center;">
								<h3 align="center"></h3>
							</td>
							<%
							End If
							%>

							<td id="fileUpload" style="display: none; text-align: center">
								<label for="filSelectFile">File:</label>
								<input type="file" name="filSelectFile" id="filSelectFile" onchange="EmbedLink()" />
							</td>

							<td width="20"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td height="10">
								<table width="100%" class="invisible cellspace0 cellpadding0">
									<tr>
										<td colspan="12"></td>
									</tr>

									<tr>
										<td></td>

										<%
											' Server / Local
											If Session("optionOLEType") < 2 Then
										%>
										<td></td>
										<td></td>
										<td width="10">
											<input id="cmdAdd" name="cmdAdd" type="button" value="Upload" class="btn" onclick="Add()"/>
										</td>
										<%
											' Linked / Embedded
										Else
										%>
										<td width="10">
											<input id="cmdLink" name="cmdLink" type="button" value="Link" class="btn hidden" onclick="showFileUpload(3)"/>
										</td>
										<td width="40">&nbsp;&nbsp;
										</td>
										<td width="10">
											<input id="cmdEmbed" name="cmdEmbed" type="button" value="Embed" class="btn" onclick="showFileUpload(2)"
											<%If Session("optionOLEMaxEmbedSize") <= 0 Then%>disabled="disabled"<%End If%>
											/>
										</td>
										<%
										End If
										%>
										<td width="40">&nbsp;&nbsp;
										</td>
										<td width="10">
											<input id="cmdEdit" name="cmdEdit" type="button" value="Download" class="btn" onclick="Edit()"/>
										</td>
										<td width="40">&nbsp;&nbsp;
										</td>

										<%
											' Server / Local
											If Session("optionOLEType") < 2 Then
												' Clear
										%>
										<td width="10">
											<input id="cmdClear" name="cmdClear" type="button" value="None" class="btn" onclick="Clear()"/>
										</td>

										<td width="40">&nbsp;&nbsp;
										</td>
										<td width="10">
											<input id="cmdSelect" name="cmdSelect" type="button" value="Select" class="btn" onclick="Select()"/>
											<input id="cmdUpload2" name="cmdUpload2" type="submit" value="OK" style="display: none;" />
										</td>
										<%
										Else
											' Properties button
										%>
										<td width="10">
											<input id="cmdProperties" name="cmdProperties" type="button" value="Properties" class="btn" onclick="Properties()"/>
										</td>
										<td width="40">
										&nbsp;&nbsp;

															<td width="10">
																<input id="cmdRemove" name="cmdRemove" type="button" value="Clear" class="btn" onclick="Remove()"/>
															</td>

										<td width="40">&nbsp;&nbsp;
										</td>		
										<td width="10">
											<input id="cmdSelect" name="cmdSelect" type="submit" value="OK" disabled="disabled" class="btn" />
										</td>								
										<%
										End If
										%>
										<td width="40">&nbsp;&nbsp;
										</td>
										<td width="10">
											<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" class="btn" onclick="Cancel()"/>
										</td>
									</tr>
								</table>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<%
			Response.Write("<INPUT type='hidden' id='txtErrorDescription' name='txtErrorDescription' value=''>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id='txtOptionColumnID' name='txtOptionColumnID' value='" & Session("optionColumnID") & "'>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id='txtFile' name='txtFile' value=""" & Replace(Session("optionFile").ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id='txtFFOLEType' name='txtFFOLEType' value='" & Session("optionOLEType") & "'>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id='txtOLEMaxEmbedSize' name='txtOLEMaxEmbedSize' value='" & Session("optionOLEMaxEmbedSize") & "'>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id='txtOLEReadOnly' name='txtOLEReadOnly' value='" & Session("optionOLEReadOnly") & "'>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id='txtIsPhoto' name='txtIsPhoto' value='" & Session("optionIsPhoto") & "'>" & vbCrLf)
			
			' Create the document from the database into the temporary UNC path
			Dim strUploadPath As String
			Dim strFullFileName As String
			Dim strJustFileName As String
			Dim bEncryption As Boolean
			Dim strFileUncPath As String
			Dim bIsNew As Boolean
			Dim strFileSize As String
			Dim strDateModified As String
			Dim bIsReadOnly As Boolean

			bIsReadOnly = False
			bIsNew = True
			bEncryption = True
			strFullFileName = ""
			strJustFileName = ""
			strFileUncPath = ""
			strUploadPath = "\\" & Request.ServerVariables("SERVER_NAME") & "\HRProTemp$\"
			strFileSize = ""
			strDateModified = ""

			If Session("optionOLEType") > 1 Then
				Dim objOLE As HR.Intranet.Server.Ole = Session("OLEObject")
				' The following are now set using getpropertiesfromstream.
				objOLE.FileName = ""
				objOLE.TempLocationPhysical = strUploadPath
				objOLE.TempLocationUNC = strUploadPath
				objOLE.CurrentSessionKey = Session.SessionID
				objOLE.CurrentUser = Request.ServerVariables("LOGON_USER")
				objOLE.UseFileSecurity = True
				objOLE.UseEncryption = bEncryption
				objOLE.UseFileSecurity = False
				objOLE.CreateOLEDocument(Session("optionRecordID"), Session("optionColumnID"), Session("realSource"))
				bEncryption = (objOLE.OLEType = 2)
				Session("optionOLEType") = objOLE.OLEType
				strFullFileName = objOLE.FileName
				If objOLE.OLEType = 3 Then
					strJustFileName = objOLE.Filename
				Else
					strJustFileName = objOLE.DisplayFilename
				End If
				
				strFileUncPath = objOLE.UNCAndPath
				strFileSize = objOLE.DocumentSize
				strDateModified = objOLE.DocumentModifyDate
				bIsNew = (Len(strJustFileName) = 0)
				Session("OLEObject") = objOLE
				objOLE = Nothing				
			End If
			
			
		%>
		<input type='hidden' id="txtOLEServerPath" name="txtOLEServerPath" value="">
		<input type='hidden' id="txtOLELocalPath" name="txtOLELocalPath" value="">
		<input type='hidden' id="txtPicturePath" name="txtPicturePath" value="">

		<input type="hidden" id="txtOLEType" name="txtOLEType" value='<%=session("optionOLEType")%>'>
		<input type="hidden" id="txtOLEFile" name="txtOLEFile" value="<%=strJustFileName%>">
		<input type="hidden" id="txtOLEFileUNCPath" name="txtOLEFileUNCPath" value="<%=strFileUncPath%>">
		<input type="hidden" id="txtOLEJustFileName" name="txtOLEJustFileName" value="<%=strJustFileName%>">
		<input type="hidden" id="txtOLEEncryption" name="txtOLEEncryption" value="<%=bEncryption%>">
		<input type="hidden" id="txtOLESession" name="txtOLESession" value="<%=session.SessionID%>">
		<input type="hidden" id="txtOLEIsNew" name="txtOLEIsNew" value="<%=bIsNew%>">
		<input type="hidden" id="txtOLECommit" name="txtOLECommit" value="0">
		<input type='hidden' id="txtOLEUploadPath" name="txtOLEUploadPath" value="<%=strUploadPath%>">
		<input type='hidden' id="txtOLEFileSize" name="txtOLEFileSize" value="<%=strFileSize%>">
		<input type='hidden' id="txtOLEModifiedDate" name="txtOLEModifiedDate" value="<%=strDateModified%>">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
		<%=Html.AntiForgeryToken()%>
	</form>

	<input type="file" id="fileAddFile" name="fileAddFile" style="height: 22px; position: absolute; top: 0; left: -9999em;">
	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<script type="text/javascript">
	oleFind_window_onload();
	$("#optionframe").dialog('open');
</script>