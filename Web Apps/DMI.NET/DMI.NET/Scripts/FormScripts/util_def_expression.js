
var frmOriginalDefinition = OpenHR.getForm("divDefExpression", "frmOriginalDefinition");
var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");
var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");


function buildjsTree() {
	var options = {};
	//BUG: why no dots/icons?
	options["themes"] = {
		"dots": true, "icons": false
		,"theme": "adv_themeroller",
		"url": window.ROOT + "Scripts/jquery/jstree/themes/adv_themeroller/style.css"
	};
	options["plugins"] = ["html_data", "ui", "contextmenu", "crrm", "hotkeys", "themes", "themeroller"];

	var hotkey = {};
	$('input[id^="txtShortcutKeys_"]').each(function () {
		var hotkeychar = $(this).val();

		$.each(hotkeychar.split(''), function (intIndex, objValue) {
			hotkey[objValue] = function () {
				SSTree1_keyPress(objValue);
			};
		});

	});

	//override default hotkeys...
	hotkey["del"] = function () {
		deleteClick();
	};
	hotkey["f2"] = function () {
		var obj = (this.data.ui.last_selected);
		if (($(obj).attr('id').substr(0, 1) == "C") || ($(obj).hasClass('root'))) return false;
		abExprMenu_Click('ID_Rename');
	}
	hotkey["ctrl+c"] = function() {
		abExprMenu_Click('ID_Copy');
	};
	hotkey["ctrl+x"] = function () {
		abExprMenu_Click('ID_Cut');
	};
	hotkey["ctrl+v"] = function () {
		abExprMenu_Click('ID_Paste');
	};
	hotkey["shift+="] = function() {
		SSTree1_keyPress("+");
	};
	hotkey["shift+,"] = function() {
		SSTree1_keyPress("<");
	};
	hotkey["shift+."] = function () {
		SSTree1_keyPress(">");
	};

	options["hotkeys"] = hotkey;

	options["contextmenu"] = { "items": customMenu };
	options["types"] = {
		"types": {
			"disabled": {
				"select_node": false,
				"open_node": false,
				"close_node": false,
				"create_node": false,
				"delete_node": false
			}
		}
	};
	options["themeroller"] = {
		"item_leaf": false,
		"item_clsd": false,
		"item_open": false,
		"item": "ui-menu-item"
	};

	//set Initial Expanded Nodes
	var tree;
	
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	switch (frmUseful.txtExprNodeMode.value) {
		case "2":
			//expand all
			tree = $("#SSTree1");
			tree.bind("loaded.jstree", function (event, data) {
				tree.jstree("open_all");
				tree_SelectRootNode();
			});
			break;
		case "4":
			// Expand Top Level.
			var topLevelNodeID = $('.root>ul>li').attr('id');
			tree = $("#SSTree1");
			// ReSharper disable once UnusedParameter
			tree.bind("loaded.jstree", function (event, data) {
				$.jstree._reference("#SSTree1").open_node('#' + tree_getRootNodeID());
				$.jstree._reference("#SSTree1").open_node('#' + topLevelNodeID);
				$('#SSTree1').jstree('refresh');
				tree_SelectRootNode();
			});
			break;
		default:
			tree = $("#SSTree1");
			tree.bind("loaded.jstree", function (event, data) {
				$('#SSTree1').jstree('refresh');
				tree_SelectRootNode();
			});
			break;
	}

	options["core"] = { 'check_callback': true };	// Must have - this enables inline renaming etc...

	//Convert the <ul><li> structure to a jsTree
	try {
		$('#SSTree1').jstree("set_theme", "apple", "/Scripts/jquery/jstree/theme/apple");
		$('#SSTree1').jstree(options);
		$("#SSTree1").bind(
						//click event
						"select_node.jstree", function (evt, data) {
							refreshControls();
						}
		);
		$('#SSTree1').bind("paste.jstree", function(event, data) {
			resetIDandTag(data.rslt);

			if (frmUseful.txtCutCopyType.value == "CUT") {

				var pastedId = data.rslt.nodes[0].id;
				$.jstree._focused().deselect_all();
				$.jstree._focused().select_node("#" + pastedId);
				//Turn into a copy now, for repeated pasting...
				frmUseful.txtCutCopyType.value = "COPY";
				$.jstree._focused().copy();

			}

		});

		$('#SSTree1').bind("dblclick.jstree", function () {
			SSTree1_dblClick();
			return false;
		});		

	}
	catch (e) {
		alert("Unable to generate expression tree.\n" + e);
	}



}



function util_def_expression_onload() {


	resizeGridToFit(); 


	$("#workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");

	var fOK = true;
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");
	var sErrMsg = frmUseful.txtErrorDescription.value;
	if (sErrMsg.length > 0) {
		fOK = false;
		OpenHR.messageBox(sErrMsg);
	}

	if (fOK == true) {

		if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
			frmUseful.txtUtilID.value = 0;
			frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
			frmDefinition.txtDescription.value = "";

			$('#SSTree1').append('<ul><li class="root" data-nodetype="root" id="E0" data-tag=""><a style="font-weight: bold;" href="#"> </a></li></ul>');

		} else {
			loadDefinition();
		}
		try {
			frmDefinition.txtName.focus();
		} catch (e) {
			
		}

		buildjsTree();
		

		frmUseful.txtLoading.value = 'N';
		try {			
			frmDefinition.txtName.focus();
		} catch (e) {
		}

		// Get menu.asp to refresh the menu.
		menu_refreshMenu();
		refreshControls();
		$('#cmdCancel').hide();

	}


}


function resizeGridToFit() {

	//resize grid	
	var workPageHeight = $('#frmDefinition>div.absolutefull').height();
	var gridTopPos = $('div.gridwithbuttons').position().top;
	var newGridHeight = workPageHeight - gridTopPos;

	$('#SSTree1').height(newGridHeight);
	$('#SSTree1').width($('div.stretchyfill').outerWidth(true));
}

function resetIDandTag(dataObj) {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	//do nothing if cutting///
	if (frmUseful.txtCutCopyType.value == "CUT") {		
		return true;
	}

	try {
		var sOldID = dataObj.nodes[0].id;
		var sNewID = getUniqueNodeKey("C");
		
		$("#copy_" + sOldID).attr('id', sNewID);

		$.jstree._focused().deselect_all();
		$.jstree._focused().select_node('#' + sNewID);

		var sOldParentID = $.jstree._focused()._get_parent('#' + sOldID).attr('id');
		var sNewParentID = $.jstree._focused()._get_parent('#' + sNewID).attr('id');

		var sOldTag = $('#' + sNewID).attr('data-tag');
		var sNewTag = sOldTag.replace(sOldID.substr(1), sNewID.substr(1)).replace(sOldParentID.substr(1), sNewParentID.substr(1));

		$('#' + sNewID).attr('data-tag', sNewTag);

		resetsubIDandTags(sNewID);

	} catch (e) {		
		alert("(resetIDandTag)\n" + e);
		return false;
	}
}

function resetsubIDandTags(parentObjID) {

	//Now update children's IDs
	$('#' + parentObjID + '>ul>li').each(function () {
		try {
		var sType = ($(this).attr('id').substr(0, 1) == 'E') ? 'C' : 'E';

		var sOldID = $(this).attr('id').substr(5);
		var sNewID = getUniqueNodeKey(sType);

		$("#copy_" + sOldID).attr('id', sNewID);

		var sOldParentID = $.jstree._focused()._get_parent('#' + sOldID).attr('id');
		var sNewParentID = $.jstree._focused()._get_parent('#' + sNewID).attr('id');

		var sOldTag = $('#' + sNewID).attr('data-tag');
		var sNewTag = sOldTag.replace(sOldID.substr(1), sNewID.substr(1)).replace(sOldParentID.substr(1), sNewParentID.substr(1));

		$('#' + sNewID).attr('data-tag', sNewTag);

		resetsubIDandTags(sNewID);

		} catch (e) {
			alert("(resetsubIDandTags)\n" + e);
			return false;
		}
	});
}

function loadDefinition() {
	var sKey;
	var frmOriginalDefinition = OpenHR.getForm("divDefExpression", "frmOriginalDefinition");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	var dataCollection = frmOriginalDefinition.elements;
	if (dataCollection != null) {
		for (var i = 0; i < dataCollection.length; i++) {
			var sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 10);
			if (sControlName == "txtDefn_E_") {
				var sExprDefn = dataCollection.item(i).value;
				if (expressionParameter(sExprDefn, "PARENTCOMPONENTID") == 0) {

					if (frmUseful.txtAction.value.toUpperCase() == "COPY") {
						frmUseful.txtUtilID.value = 0;
						frmDefinition.txtName.value = "Copy of " + expressionParameter(sExprDefn, "NAME");
						frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
						frmUseful.txtChanged.value = 1;
					}
					else {
						frmDefinition.txtName.value = expressionParameter(sExprDefn, "NAME");
						frmDefinition.txtOwner.value = expressionParameter(sExprDefn, "USERNAME");
					}

					frmDefinition.txtDescription.value = expressionParameter(sExprDefn, "DESCRIPTION");

					var sAccess = expressionParameter(sExprDefn, "ACCESS");
					if (sAccess == "RW") {
						frmDefinition.optAccessRW.checked = true;
					}
					else {
						if (sAccess == "RO") {
							frmDefinition.optAccessRO.checked = true;
						}
						else {
							frmDefinition.optAccessHD.checked = true;
						}
					}
					frmOriginalDefinition.txtOriginalAccess.value = sAccess;


					sKey = "E" + expressionParameter(sExprDefn, "EXPRID");

					//Add the title node
					$('#SSTree1').append('<ul><li class="root" data-nodetype="root" id="' + sKey + '"><a style="font-weight: bold;" href="#"> </a></li></ul>');

					$('#' + sKey).attr('data-tag', sExprDefn);
					// Load the expression definition into the treeview.
					loadComponentNodes(expressionParameter(sExprDefn, "EXPRID"), true);

					break;
				}
			}
		}
	}

	// If its read only, disable everything.
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {	
		setTimeout('expr_disableAllProps()', 100);
	
		button_disable(frmDefinition.cmdPrint, false);
		if (frmUseful.txtUtilType.value == 11) {
			button_disable(frmDefinition.cmdTest, false);
		}
	}
}

function expr_disableAllProps() {
	$('#frmDefinition #nav input, #frmDefinition #nav textarea').prop('disabled', true);
}

function loadComponentNodes(piExprID, pfVisible) {
	var i;
	var sParentKey = "E" + piExprID;
	var sControlName;
	var sComponentDefn;

	var frmOriginalDefinition = OpenHR.getForm("divDefExpression", "frmOriginalDefinition");

	var dataCollection = frmOriginalDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 10);
			if (sControlName == "txtDefn_C_") {
				sComponentDefn = dataCollection.item(i).value;
				if (componentParameter(sComponentDefn, "EXPRID") == piExprID) {
					/* Load node and then load sub-expressions */

					//add ul to parent if missing
					if ($('#' + sParentKey).children('ul').length == 0) {
						$('#' + sParentKey).append('<ul></ul>');
					}

					//append new node to parent
					var sText = componentDescription(sComponentDefn).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
					var nodeID = "C" + componentParameter(sComponentDefn, "COMPONENTID");
					$('#' + sParentKey + ' ul').first().append('<li id="' + nodeID + '"><a style="font-weight: normal;" href="#">' + sText + '</a></li>');
					$('#' + nodeID).attr('data-tag', sComponentDefn);

					//Set colour of node.
					$('#' + nodeID + ' a').css('color', getNodeColour(tree_SelectedItemLevel("#" + nodeID)));

					loadSubExpressionsNodes(componentParameter(sComponentDefn, "COMPONENTID"), true);
				}
			}
		}
	}
}

function loadSubExpressionsNodes(piComponentID, pfVisible) {
	var i;
	var sControlName;
	var sExprDefn;

	var sParentKey = "C" + piComponentID;
	var frmOriginalDefinition = OpenHR.getForm("divDefExpression", "frmOriginalDefinition");

	var dataCollection = frmOriginalDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 10);
			if (sControlName == "txtDefn_E_") {
				sExprDefn = dataCollection.item(i).value;
				if (expressionParameter(sExprDefn, "PARENTCOMPONENTID") == piComponentID) {
					/* Load node and then load components */

					//add ul to parent if missing
					if ($('#' + sParentKey).children('ul').length == 0) {
						$('#' + sParentKey).append('<ul></ul>');
					}

					//append new node to parent
					var sText = expressionParameter(sExprDefn, "NAME").replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
					var nodeID = "E" + expressionParameter(sExprDefn, "EXPRID");
					$('#' + sParentKey + ' ul').first().append('<li id="' + nodeID + '"><a style="font-weight: normal;" href="#">' + sText + '</a></li>');
					$('#' + nodeID).attr('data-tag', sExprDefn);

					//Set colour of node.
					$('#' + nodeID + ' a').css('color', getNodeColour(tree_SelectedItemLevel("#" + nodeID)));

					loadComponentNodes(expressionParameter(sExprDefn, "EXPRID"), true);
				}
			}
		}
	}
}

function getNodeColour(piLevel) {
	var sColour;
	var iModLevel;
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	sColour = 'rgb(0, 0, 0)'; //6697779;

	if (frmUseful.txtExprColourMode.value == 2) {
		iModLevel = piLevel % 7;

		switch (iModLevel) {
			case 0:
				sColour = 'rgb(105,105,129)'; //13111040;
				break;
			case 1:
				sColour = 'rgb(0,15,200)'; //0;
				break;
			case 2:
				sColour = 'rgb(192,0,54)'; //180;
				break;
			case 3:
				sColour = 'rgb(56,125,54)'; //32000;
				break;
			case 4:
				sColour = 'rgb(56,0,145)'; //8192000;
				break;
			case 5:
				sColour = 'rgb(125,125,0)'; //32125;
				break;
			case 6:
				sColour = 'rgb(0,125,125)'; //8224000;
				break;
			default:
				sColour = 'rgb(105,105,129)'; //8192125;
		}
	}
	return sColour;
}

function expressionParameter(psDefnString, psParameter) {
	var iCharIndex;
	var sDefn;

	sDefn = new String(psDefnString);

	iCharIndex = sDefn.indexOf("	");
	if (iCharIndex >= 0) {
		if (psParameter == "EXPRID") return sDefn.substr(0, iCharIndex);
		sDefn = sDefn.substr(iCharIndex + 1);
		iCharIndex = sDefn.indexOf("	");
		if (iCharIndex >= 0) {
			if (psParameter == "NAME") return sDefn.substr(0, iCharIndex);
			sDefn = sDefn.substr(iCharIndex + 1);
			iCharIndex = sDefn.indexOf("	");
			if (iCharIndex >= 0) {
				if (psParameter == "TABLEID") return sDefn.substr(0, iCharIndex);
				sDefn = sDefn.substr(iCharIndex + 1);
				iCharIndex = sDefn.indexOf("	");
				if (iCharIndex >= 0) {
					if (psParameter == "RETURNTYPE") return sDefn.substr(0, iCharIndex);
					sDefn = sDefn.substr(iCharIndex + 1);
					iCharIndex = sDefn.indexOf("	");
					if (iCharIndex >= 0) {
						if (psParameter == "RETURNSIZE") return sDefn.substr(0, iCharIndex);
						sDefn = sDefn.substr(iCharIndex + 1);
						iCharIndex = sDefn.indexOf("	");
						if (iCharIndex >= 0) {
							if (psParameter == "RETURNDECIMALS") return sDefn.substr(0, iCharIndex);
							sDefn = sDefn.substr(iCharIndex + 1);
							iCharIndex = sDefn.indexOf("	");
							if (iCharIndex >= 0) {
								if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
								sDefn = sDefn.substr(iCharIndex + 1);
								iCharIndex = sDefn.indexOf("	");
								if (iCharIndex >= 0) {
									if (psParameter == "PARENTCOMPONENTID") return sDefn.substr(0, iCharIndex);
									sDefn = sDefn.substr(iCharIndex + 1);
									iCharIndex = sDefn.indexOf("	");
									if (iCharIndex >= 0) {
										if (psParameter == "USERNAME") return sDefn.substr(0, iCharIndex);
										sDefn = sDefn.substr(iCharIndex + 1);
										iCharIndex = sDefn.indexOf("	");
										if (iCharIndex >= 0) {
											if (psParameter == "ACCESS") return sDefn.substr(0, iCharIndex);
											sDefn = sDefn.substr(iCharIndex + 1);
											iCharIndex = sDefn.indexOf("	");
											if (iCharIndex >= 0) {
												if (psParameter == "DESCRIPTION") return sDefn.substr(0, iCharIndex);
												sDefn = sDefn.substr(iCharIndex + 1);
												iCharIndex = sDefn.indexOf("	");
												if (iCharIndex >= 0) {
													if (psParameter == "TIMESTAMP") return sDefn.substr(0, iCharIndex);
													sDefn = sDefn.substr(iCharIndex + 1);
													iCharIndex = sDefn.indexOf("	");
													if (iCharIndex >= 0) {
														if (psParameter == "VIEWINCOLOUR") return sDefn.substr(0, iCharIndex);
														sDefn = sDefn.substr(iCharIndex + 1);
														if (psParameter == "EXPANDEDNODE") return sDefn;
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
	}

	return "";
}

function componentParameter(psDefnString, psParameter) {
	var iCharIndex;
	var sDefn;

	sDefn = new String(psDefnString);

	iCharIndex = sDefn.indexOf("	");
	if (iCharIndex >= 0) {
		if (psParameter == "COMPONENTID") return sDefn.substr(0, iCharIndex);
		sDefn = sDefn.substr(iCharIndex + 1);
		iCharIndex = sDefn.indexOf("	");
		if (iCharIndex >= 0) {
			if (psParameter == "EXPRID") return sDefn.substr(0, iCharIndex);
			sDefn = sDefn.substr(iCharIndex + 1);
			iCharIndex = sDefn.indexOf("	");
			if (iCharIndex >= 0) {
				if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
				sDefn = sDefn.substr(iCharIndex + 1);
				iCharIndex = sDefn.indexOf("	");
				if (iCharIndex >= 0) {
					if (psParameter == "FIELDCOLUMNID") return sDefn.substr(0, iCharIndex);
					sDefn = sDefn.substr(iCharIndex + 1);
					iCharIndex = sDefn.indexOf("	");
					if (iCharIndex >= 0) {
						if (psParameter == "FIELDPASSBY") return sDefn.substr(0, iCharIndex);
						sDefn = sDefn.substr(iCharIndex + 1);
						iCharIndex = sDefn.indexOf("	");
						if (iCharIndex >= 0) {
							if (psParameter == "FIELDSELECTIONTABLEID") return sDefn.substr(0, iCharIndex);
							sDefn = sDefn.substr(iCharIndex + 1);
							iCharIndex = sDefn.indexOf("	");
							if (iCharIndex >= 0) {
								if (psParameter == "FIELDSELECTIONRECORD") return sDefn.substr(0, iCharIndex);
								sDefn = sDefn.substr(iCharIndex + 1);
								iCharIndex = sDefn.indexOf("	");
								if (iCharIndex >= 0) {
									if (psParameter == "FIELDSELECTIONLINE") return sDefn.substr(0, iCharIndex);
									sDefn = sDefn.substr(iCharIndex + 1);
									iCharIndex = sDefn.indexOf("	");
									if (iCharIndex >= 0) {
										if (psParameter == "FIELDSELECTIONORDERID") return sDefn.substr(0, iCharIndex);
										sDefn = sDefn.substr(iCharIndex + 1);
										iCharIndex = sDefn.indexOf("	");
										if (iCharIndex >= 0) {
											if (psParameter == "FIELDSELECTIONFILTER") return sDefn.substr(0, iCharIndex);
											sDefn = sDefn.substr(iCharIndex + 1);
											iCharIndex = sDefn.indexOf("	");
											if (iCharIndex >= 0) {
												if (psParameter == "FUNCTIONID") return sDefn.substr(0, iCharIndex);
												sDefn = sDefn.substr(iCharIndex + 1);
												iCharIndex = sDefn.indexOf("	");
												if (iCharIndex >= 0) {
													if (psParameter == "CALCULATIONID") return sDefn.substr(0, iCharIndex);
													sDefn = sDefn.substr(iCharIndex + 1);
													iCharIndex = sDefn.indexOf("	");
													if (iCharIndex >= 0) {
														if (psParameter == "OPERATORID") return sDefn.substr(0, iCharIndex);
														sDefn = sDefn.substr(iCharIndex + 1);
														iCharIndex = sDefn.indexOf("	");
														if (iCharIndex >= 0) {
															if (psParameter == "VALUETYPE") return sDefn.substr(0, iCharIndex);
															sDefn = sDefn.substr(iCharIndex + 1);
															iCharIndex = sDefn.indexOf("	");
															if (iCharIndex >= 0) {
																if (psParameter == "VALUECHARACTER") return sDefn.substr(0, iCharIndex);
																sDefn = sDefn.substr(iCharIndex + 1);


																iCharIndex = sDefn.indexOf("	");
																if (iCharIndex >= 0) {
																	if (psParameter == "VALUENUMERIC") return sDefn.substr(0, iCharIndex);
																	sDefn = sDefn.substr(iCharIndex + 1);
																	iCharIndex = sDefn.indexOf("	");
																	if (iCharIndex >= 0) {
																		if (psParameter == "VALUELOGIC") return sDefn.substr(0, iCharIndex);
																		sDefn = sDefn.substr(iCharIndex + 1);
																		iCharIndex = sDefn.indexOf("	");
																		if (iCharIndex >= 0) {
																			if (psParameter == "VALUEDATE") return sDefn.substr(0, iCharIndex);
																			sDefn = sDefn.substr(iCharIndex + 1);
																			iCharIndex = sDefn.indexOf("	");
																			if (iCharIndex >= 0) {
																				if (psParameter == "PROMPTDESCRIPTION") return sDefn.substr(0, iCharIndex);
																				sDefn = sDefn.substr(iCharIndex + 1);
																				iCharIndex = sDefn.indexOf("	");
																				if (iCharIndex >= 0) {
																					if (psParameter == "PROMPTMASK") return sDefn.substr(0, iCharIndex);
																					sDefn = sDefn.substr(iCharIndex + 1);
																					iCharIndex = sDefn.indexOf("	");
																					if (iCharIndex >= 0) {
																						if (psParameter == "PROMPTSIZE") return sDefn.substr(0, iCharIndex);
																						sDefn = sDefn.substr(iCharIndex + 1);
																						iCharIndex = sDefn.indexOf("	");
																						if (iCharIndex >= 0) {
																							if (psParameter == "PROMPTDECIMALS") return sDefn.substr(0, iCharIndex);
																							sDefn = sDefn.substr(iCharIndex + 1);
																							iCharIndex = sDefn.indexOf("	");
																							if (iCharIndex >= 0) {
																								if (psParameter == "FUNCTIONRETURNTYPE") return sDefn.substr(0, iCharIndex);
																								sDefn = sDefn.substr(iCharIndex + 1);
																								iCharIndex = sDefn.indexOf("	");
																								if (iCharIndex >= 0) {
																									if (psParameter == "LOOKUPTABLEID") return sDefn.substr(0, iCharIndex);
																									sDefn = sDefn.substr(iCharIndex + 1);
																									iCharIndex = sDefn.indexOf("	");
																									if (iCharIndex >= 0) {
																										if (psParameter == "LOOKUPCOLUMNID") return sDefn.substr(0, iCharIndex);
																										sDefn = sDefn.substr(iCharIndex + 1);
																										iCharIndex = sDefn.indexOf("	");
																										if (iCharIndex >= 0) {
																											if (psParameter == "FILTERID") return sDefn.substr(0, iCharIndex);
																											sDefn = sDefn.substr(iCharIndex + 1);
																											iCharIndex = sDefn.indexOf("	");
																											if (iCharIndex >= 0) {
																												if (psParameter == "EXPANDEDNODE") return sDefn.substr(0, iCharIndex);
																												sDefn = sDefn.substr(iCharIndex + 1);
																												iCharIndex = sDefn.indexOf("	");
																												if (iCharIndex >= 0) {
																													if (psParameter == "PROMPTDATETYPE") return sDefn.substr(0, iCharIndex);
																													sDefn = sDefn.substr(iCharIndex + 1);
																													iCharIndex = sDefn.indexOf("	");
																													if (iCharIndex >= 0) {
																														if (psParameter == "DESCRIPTION") return sDefn.substr(0, iCharIndex);
																														sDefn = sDefn.substr(iCharIndex + 1);
																														iCharIndex = sDefn.indexOf("	");
																														if (iCharIndex >= 0) {
																															if (psParameter == "FIELDTABLEID") return sDefn.substr(0,

																																	iCharIndex);
																															sDefn = sDefn.substr(iCharIndex + 1);
																															iCharIndex = sDefn.indexOf("	");
																															if (iCharIndex >= 0) {
																																if (psParameter == "FIELDSELECTIONORDERNAME") return false;

																																sDefn.substr(0, iCharIndex);
																																sDefn = sDefn.substr(iCharIndex + 1);





																																if (psParameter == "FIELDSELECTIONFILTERNAME") return sDefn;
																															}
																														}
																													}
																												}
																											}
																										}
																									}
																								}
																							}
																						}
																					}
																				}
																			}
																		}
																	}
																}
															}
														}
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
	}

	return "";
}

function componentDescription(psDefnString) {
	var sDesc;
	var reDecimalSeparator = new RegExp("\\.", "gi");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	sDesc = "";

	if ((componentParameter(psDefnString, "TYPE") == "4") ||
			(componentParameter(psDefnString, "TYPE") == "6")) {
		// Value or Lookup Value.
		switch (componentParameter(psDefnString, "VALUETYPE")) {
			case "1":
				// Character value.
				sDesc = "\"" + componentParameter(psDefnString, "VALUECHARACTER") + "\"";
				break;

			case "2":
				// Numeric value.				
				sDesc = componentParameter(psDefnString, "VALUENUMERIC");
				sDesc = parseFloat(sDesc).toString();
				sDesc = sDesc.replace(reDecimalSeparator, frmUseful.txtLocaleDecimal.value);
				break;

			case "3":
				// Logic value.
				if (componentParameter(psDefnString, "VALUELOGIC") == "1") {
					sDesc = "True";
				}
				else {
					sDesc = "False";
				}
				break;

			case "4":
				// Date value.
				sDesc = componentParameter(psDefnString, "VALUEDATE");
				if (sDesc.length == 0) {
					sDesc = "Empty Date";
				}
				else {
					sDesc = OpenHR.ConvertSQLDateToLocale(sDesc);
				}
		}
	}
	else {
		if (componentParameter(psDefnString, "TYPE") == "7") {
			// Prompted Value.
			sDesc = componentParameter(psDefnString, "PROMPTDESCRIPTION") + " : ";

			switch (componentParameter(psDefnString, "VALUETYPE")) {
				case "1":
					// Character value.
					sDesc = sDesc + "<string>";
					break;

				case "2":
					// Numeric value.
					sDesc = sDesc + "<numeric>";
					break;

				case "3":
					// Logic value.
					sDesc = sDesc + "<logic>";
					break;

				case "4":
					// Date value.
					sDesc = sDesc + "<date>";
					break;

				case "5":
					// lookup value.
					sDesc = sDesc + "<lookup value>";
			}
		}
		else {
			sDesc = componentParameter(psDefnString, "DESCRIPTION");
		}
	}

	return sDesc;
}

function refreshControls() {	
	var sKey;
	var fViewing;
	var fIsNotOwner;
	var fDisableAdd;
	var fDisableEdit;
	var fDisableDelete;
	var fDisableInsert;
	var fDisableCut;
	var fDisableCopy;
	var fDisablePaste;
	var fDisableMoveDown;
	var fDisableMoveUp;
	var iNodesSelected;

	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	fViewing = (frmUseful.txtAction.value.toUpperCase() == "VIEW");
	fIsNotOwner = (frmUseful.txtUserName.value.toUpperCase() != frmDefinition.txtOwner.value.toUpperCase());
	
	radio_disable(frmDefinition.optAccessRW, ((fIsNotOwner) || (fViewing)));
	radio_disable(frmDefinition.optAccessRO, ((fIsNotOwner) || (fViewing)));
	radio_disable(frmDefinition.optAccessHD, ((fIsNotOwner) || (fViewing)));

	fDisableAdd = fViewing;
	fDisableEdit = fViewing;
	fDisableDelete = fViewing;
	fDisableInsert = fViewing;
	fDisableCut = fViewing;
	fDisableCopy = fViewing;
	fDisablePaste = fViewing;
	fDisableMoveDown = fViewing;
	fDisableMoveUp = fViewing;
	iNodesSelected = 0;

	if (tree_SelectedItemKey() == undefined) {
		// Select the root node.		
		$("#SSTree1").jstree("select_node", ".root");
	}

	// Loop through each selected node
	$('#SSTree1 .jstree-clicked').each(function () {
		iNodesSelected = iNodesSelected + 1;

		if (tree_SelectedItemLevel('#' + $(this).parent().attr('id')) == 1) {
			// If the root node is selected then disable the Insert/Modify/Delete buttons.
			fDisableInsert = true;
			fDisableEdit = true;
			fDisableDelete = true;
			fDisableCut = true;
			fDisableCopy = true;
			fDisableMoveDown = true;
			fDisableMoveUp = true;
		}
		else {
			sKey = $(this).parent().attr('id');

			if (sKey.substr(0, 1) == "E") {
				fDisableEdit = true;
				fDisableInsert = true;
				fDisableDelete = true;
				fDisableCut = true;
				fDisableCopy = true;
				fDisableMoveDown = true;
				fDisableMoveUp = true;
			}
			else {
				if (tree_LastSiblingID() == tree_selectedNodeID()) {
					fDisableMoveDown = true;
				}

				if (tree_FirstSiblingID() == tree_selectedNodeID()) {
					fDisableMoveUp = true;
				}
			}
		}
		//}
	});

	// Only allow edit and insert when single nodes are selected
	if (iNodesSelected != 1) {
		fDisableInsert = true;
		fDisableEdit = true;
		fDisableMoveDown = true;
		fDisableMoveUp = true;
	}

	if (iNodesSelected == 0) {
		fDisableDelete = true;
	}

	if ((frmUseful.txtCutCopyType.value != "COPY") && (frmUseful.txtCutCopyType.value != "CUT")) {
		fDisablePaste = true;
	}

	// Enable/disable controls depending on the selected component.
	button_disable(frmDefinition.cmdAdd, fDisableAdd);
	button_disable(frmDefinition.cmdInsert, fDisableInsert);
	button_disable(frmDefinition.cmdEdit, fDisableEdit);
	button_disable(frmDefinition.cmdDelete, fDisableDelete);
	//button_disable(frmDefinition.cmdPrint, true);

	//button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
	//		(fViewing == true)));

	menu_toolbarEnableItem('mnutoolSaveReport', (!((frmUseful.txtChanged.value == 0) ||
			(fViewing == true))));

	if (fDisableMoveDown == true) {
		frmUseful.txtCanMoveDown.value = 0;
	}
	else {
		frmUseful.txtCanMoveDown.value = 1;
	}

	if (fDisableMoveUp == true) {
		frmUseful.txtCanMoveUp.value = 0;
	}
	else {
		frmUseful.txtCanMoveUp.value = 1;
	}

	if (fDisableCopy == true) {
		frmUseful.txtCanCopy.value = 0;
	}
	else {
		frmUseful.txtCanCopy.value = 1;
	}

	if (fDisablePaste == true) {
		frmUseful.txtCanPaste.value = 0;
	}
	else {
		frmUseful.txtCanPaste.value = 1;
	}

	if (fDisableCut == true) {
		frmUseful.txtCanCut.value = 0;
	}
	else {
		frmUseful.txtCanCut.value = 1;
	}
}

function changeName() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	$('#SSTree1').jstree('rename_node', '.root', frmDefinition.txtName.value);
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function changeDescription() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function changeAccess() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function addClick() {
	var fOK;
	var sKey;
	var sRelativeKey;

	var frmOptionArea = OpenHR.getForm("optionframeset", "frmGotoOption");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	var iFunctionID = 0;
	var iParamIndex = 0;
	var nodParameter;

	fOK = true;

	frmOptionArea.txtGotoOptionAction.value = "ADDEXPRCOMPONENT";
	frmOptionArea.txtGotoOptionTableID.value = frmUseful.txtTableID.value;
	frmOptionArea.txtGotoOptionExprID.value = frmUseful.txtUtilID.value;

	sKey = tree_SelectedItemKey();	
	if (sKey.substr(0, 1) == "E") {
		sRelativeKey = sKey;
		nodParameter = $('#' + sRelativeKey);
	}
	else {
		sRelativeKey = tree_SelectedItemParentKey();
		nodParameter = $('#' + sRelativeKey);
	}
	frmOptionArea.txtGotoOptionLinkRecordID.value = sRelativeKey;

	if ((sRelativeKey.substr(0, 1) == "E") &&
			(tree_SelectedItemLevel('#' + sRelativeKey) > 1)) {

		var iType = componentParameter(nodParameter.parent().parent().attr('data-tag'), "TYPE");
		if (iType == 2) {
			// Function parameter
			iFunctionID = componentParameter(nodParameter.parent().parent().attr('data-tag'), "FUNCTIONID");

			var iLoop = 0;
			nodParameter.parent().children().each(function () {
				var nodTemp = $(this);
				iLoop += 1;
				if (nodTemp.attr('id') == nodParameter.attr('id')) {
					iParamIndex = iLoop;
					return false;
				}
			});
		}
	}

	frmOptionArea.txtGotoOptionFunctionID.value = iFunctionID;
	frmOptionArea.txtGotoOptionParameterIndex.value = iParamIndex;

	switch (frmUseful.txtUtilType.value) {
		case "11":
			// Filter
			frmOptionArea.txtGotoOptionExprType.value = 11;
			break;
		case "12":
			// Calculation
			frmOptionArea.txtGotoOptionExprType.value = 10;
			break;
		default:
			fOK = false;
	}

	if (fOK == true) {
		OpenHR.submitForm(frmOptionArea, "optionframe", null, null, "expression_addClick");
	}
}

function insertClick() {
	var fOK;
	var frmOptionArea = OpenHR.getForm("optionframeset", "frmGotoOption");
	var frmRefresh = OpenHR.getForm("refreshframe", "frmRefresh");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	var iFunctionID = 0;
	var iParamIndex = 0;

	fOK = true;
	OpenHR.submitForm(frmRefresh);

	frmOptionArea.txtGotoOptionPage.value = "util_def_exprComponent";
	frmOptionArea.txtGotoOptionAction.value = "INSERTEXPRCOMPONENT";
	frmOptionArea.txtGotoOptionTableID.value = frmUseful.txtTableID.value;
	frmOptionArea.txtGotoOptionLinkRecordID.value = tree_SelectedItemKey();
	frmOptionArea.txtGotoOptionExprID.value = frmUseful.txtUtilID.value;

	var sRelativeKey;
	var nodParameter;
	var iType;

	var sKey = tree_SelectedItemKey();
	if (sKey.substr(0, 1) == "E") {
		sRelativeKey = sKey;
		nodParameter = $('#' + sRelativeKey);
	}
	else {
		sRelativeKey = tree_SelectedItemParentKey();
		nodParameter = $('#' + sRelativeKey);
	}

	if ((sRelativeKey.substr(0, 1) == "E") &&
			(tree_SelectedItemLevel('#' + sRelativeKey) > 1)) {
		iType = componentParameter(nodParameter.parent().parent().attr('data-tag'), "TYPE");
		if (iType == 2) {
			// Function parameter
			iFunctionID = componentParameter(nodParameter.parent().parent().attr('data-tag'), "FUNCTIONID");

			var iLoop = 0;
			nodParameter.parent().children().each(function () {
				var  nodTemp = $(this);
				iLoop += 1;
				if (nodTemp.attr('id') == nodParameter.attr('id')) {
					iParamIndex = iLoop;
					return false;
				}
			});
		}
	}
	frmOptionArea.txtGotoOptionFunctionID.value = iFunctionID;
	frmOptionArea.txtGotoOptionParameterIndex.value = iParamIndex;

	switch (frmUseful.txtUtilType.value) {
		case "11":
			// Filter
			frmOptionArea.txtGotoOptionExprType.value = 11;
			break;
		case "12":
			// Calculation
			frmOptionArea.txtGotoOptionExprType.value = 10;
			break;
		default:
			fOK = false;
	}

	if (fOK == true) {
		OpenHR.submitForm(frmOptionArea, "optionframe", null, null, "expression_insertClick");
	}
}

function editClick() {	
	var fOK;
	var frmOptionArea = OpenHR.getForm("optionframeset", "frmGotoOption");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	var iFunctionID = 0;
	var iParamIndex = 0;

	fOK = true;

	frmOptionArea.txtGotoOptionAction.value = "EDITEXPRCOMPONENT";
	frmOptionArea.txtGotoOptionTableID.value = frmUseful.txtTableID.value;
	frmOptionArea.txtGotoOptionLinkRecordID.value = tree_SelectedItemKey();
	frmOptionArea.txtGotoOptionExprID.value = frmUseful.txtUtilID.value;
	frmOptionArea.txtGotoOptionExtension.value = tree_SelectedItemTag();
	var sRelativeKey;
	var nodParameter;
	var iType;
	var nodTemp;

	var sKey = tree_SelectedItemKey();
	if (sKey.substr(0, 1) == "E") {
		sRelativeKey = sKey;
		nodParameter = $('#' + sRelativeKey);
	}
	else {
		sRelativeKey = tree_SelectedItemParentKey();
		nodParameter = $('#' + sRelativeKey);
	}

	if ((sRelativeKey.substr(0, 1) == "E") &&
			(tree_SelectedItemLevel('#' + sRelativeKey) > 1)) {
		iType = componentParameter(nodParameter.parent().parent().attr('data-tag'), "TYPE");
		if (iType == 2) {
			// Function parameter
			iFunctionID = componentParameter(nodParameter.parent().parent().attr('data-tag'), "FUNCTIONID");

			var iLoop = 0;
			nodParameter.parent().children().each(function () {
				nodTemp = $(this);
				iLoop += 1;
				if (nodTemp.attr('id') == nodParameter.attr('id')) {
					iParamIndex = iLoop;
					return false;
				}
			});
		}
	}
	frmOptionArea.txtGotoOptionFunctionID.value = iFunctionID;
	frmOptionArea.txtGotoOptionParameterIndex.value = iParamIndex;

	switch (frmUseful.txtUtilType.value) {
		case "11":
			// Filter
			frmOptionArea.txtGotoOptionExprType.value = 11;
			break;
		case "12":
			// Calculation
			frmOptionArea.txtGotoOptionExprType.value = 10;
			break;
		default:
			fOK = false;
	}

	if (fOK == true) {
		OpenHR.submitForm(frmOptionArea, "optionframe", null, null, "expression_insertClick");
	}
}

function setComponent(psComponentDefn, psAction, psLinkComponentID, psFunctionParameters) {	

	var iIndex;
	var fNodeExists = false;
	var objNode;
	var sNewKey;
	var sExprName;
	var sTemp;

	$("#optionframe").attr('style', 'display: none;');
	$("#workframe").show();

	if ($('#SSTree1').find('#' + psLinkComponentID).length > 0) fNodeExists = true;

	if (fNodeExists == true) {
		if (psAction == "EDITEXPRCOMPONENT") {
			createUndoView("EDIT");
			// Add the component node for the new component definition.	
			sNewKey = getUniqueNodeKey("C");

			objNode = tree_NodesAdd(psLinkComponentID, 3, sNewKey, componentDescription(psComponentDefn), psComponentDefn);

			if (componentParameter(psComponentDefn, "TYPE") == 2) {
				// Function component. Add the parameter nodes.
				sTemp = psFunctionParameters;
				while (sTemp.length > 0) {
					iIndex = sTemp.indexOf("	");
					if (iIndex >= 0) {
						sExprName = sTemp.substr(0, iIndex);
						sTemp = sTemp.substr(iIndex + 1);
					}
					else {
						sExprName = sTemp;
						sTemp = "";
					}

					objNode = tree_NodesAdd(sNewKey, 4, getUniqueNodeKey("E"), sExprName, "													");
				}
			}

			// Remove the component node for the old component definition.	
			tree_NodesRemove(psLinkComponentID);

			$("#SSTree1").jstree("deselect_all");
			tree_Refresh();
			$("#SSTree1").jstree("select_node", "#" + sNewKey);
			tree_ExpandNode($('#' + sNewKey));
		}

		if (psAction == "ADDEXPRCOMPONENT") {
			createUndoView("ADD");

			// Add the component node for the new component definition.	
			sNewKey = getUniqueNodeKey("C");
			objNode = tree_NodesAdd(psLinkComponentID, 4, sNewKey, componentDescription(psComponentDefn), psComponentDefn);

			if (componentParameter(psComponentDefn, "TYPE") == 2) {
				// Function component. Add the parameter nodes.
				sTemp = psFunctionParameters;
				while (sTemp.length > 0) {
					iIndex = sTemp.indexOf("	");
					if (iIndex >= 0) {
						sExprName = sTemp.substr(0, iIndex);
						sTemp = sTemp.substr(iIndex + 1);
					}
					else {
						sExprName = sTemp;
						sTemp = "";
					}

					objNode = tree_NodesAdd(sNewKey, 4, getUniqueNodeKey("E"), sExprName, "													");

				}
			}

			$("#SSTree1").jstree("deselect_all");
			tree_Refresh();
			$("#SSTree1").jstree("select_node", "#" + sNewKey);
			tree_ExpandNode($('#' + sNewKey));

		}

		if (psAction == "INSERTEXPRCOMPONENT") {
			createUndoView("INSERT");
			// Add the component node for the new component definition.	
			sNewKey = getUniqueNodeKey("C");
			objNode = tree_NodesAdd(psLinkComponentID, 3, sNewKey, componentDescription(psComponentDefn), psComponentDefn);

			if (componentParameter(psComponentDefn, "TYPE") == 2) {
				// Function component. Add the parameter nodes.
				sTemp = psFunctionParameters;
				while (sTemp.length > 0) {
					iIndex = sTemp.indexOf("	");
					if (iIndex >= 0) {
						sExprName = sTemp.substr(0, iIndex);
						sTemp = sTemp.substr(iIndex + 1);
					}
					else {
						sExprName = sTemp;
						sTemp = "";
					}
					objNode = tree_NodesAdd(sNewKey, 4, getUniqueNodeKey("E"), sExprName, "													");
				}
			}

			$("#SSTree1").jstree("deselect_all");
			tree_Refresh();
			$("#SSTree1").jstree("select_node", "#" + sNewKey);
			tree_ExpandNode($('#' + sNewKey));
		}
	}

	refreshControls();
	
}

function cancelComponent() {
	// Expand the work frame and hide the option frame.
	//$("#optionframe").hide();
	//$("#workframe").show();
	try {
		$('#optionframe').dialog('close');
	} catch (e) {	};

	//frmDefinition.SSTree1.style.visibility = "visible";
	//frmDefinition.SSTree1.Refresh();
	
	menu_refreshMenu();

	//frmDefinition.SSTree1.focus();
	refreshControls();

}

function getUniqueNodeKey(psType) {	
	var sKey;
	var sNodeKey;
	var sKeyID;
	var iKeyID;
	var iMaxKeyID = 0;

	$('#SSTree1').find('li').each(function () {
		sNodeKey = $(this).attr('id');

		sKeyID = sNodeKey.substr(1);
		iKeyID = Number(sKeyID);

		if (iKeyID > iMaxKeyID) {
			iMaxKeyID = iKeyID;
		}

	});

	sKeyID = String(iMaxKeyID + 1);
	sKey = psType + sKeyID;

	return (sKey);
}

function deleteClick() {	
	// Delete the selected tree nodes.
	if (tree_selectedNodeID().substr(0, 1) != "E") {
		createUndoView("DELETE");

		$.each($('#SSTree1').jstree('get_selected'), function(obj) {
			$('#SSTree1').jstree('remove', '#' + this.id);
		});

		$('#SSTree1').jstree('deselect_all');

		refreshControls();
	}
}

function printClick(pfToPrinter) {

	var cssObj = new Array;
	cssObj.push("body {font-family: segoe ui; verdana;}");
	cssObj.push("a {text-decoration: none;}");
	cssObj.push("ins {display: none;}");
	OpenHR.printDiv('SSTree1', cssObj);
}

function printNode(pobjNode, pfToPrinter) {
	
}

function testClick() {
	var iLoop;
	var sKey;
	var sTag;
	var iType;
	var sPrompts;
	var sPromptDateType;
	var sFiltersAndCalcs;
	var sURL;
	
	var frmSend = OpenHR.getForm("divDefExpression", "frmSend");	
	var frmTest = OpenHR.getForm("divDefExpression", "frmTest");

	if (validateExpression() == false) return;
	if (populateSendForm() == false) return;

	// Create a tab-delimuted string of the prompted value definitions.
	sPrompts = "";
	sFiltersAndCalcs = "";

	//for (iLoop = 1; iLoop <= frmDefinition.SSTree1.Nodes.Count; iLoop++) {
	
	$('#SSTree1 li').each(function () {		
		sKey = $(this).attr('id');
		sTag = $(this).attr('data-tag');

		if (sKey.substr(0, 1) != "E") {
			iType = componentParameter(sTag, "TYPE");

			if (iType == 7) {
				// Construct a string of prompted value components
				sPrompts = sPrompts + sKey + "	";
				sPrompts = sPrompts + componentParameter(sTag, "PROMPTDESCRIPTION") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "VALUETYPE") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "PROMPTSIZE") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "PROMPTDECIMALS") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "PROMPTMASK") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "FIELDTABLEID") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "FIELDCOLUMNID") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "VALUECHARACTER") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "VALUENUMERIC") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "VALUELOGIC") + "	";
				sPrompts = sPrompts + componentParameter(sTag, "VALUEDATE") + "	";

				sPromptDateType = new String(componentParameter(sTag, "PROMPTDATETYPE"));
				if (sPromptDateType.length == 0) {
					sPromptDateType = "0";
				}
				sPrompts = sPrompts + sPromptDateType + "	";
			}

			if (iType == 10) {
				// Filter (might include prompts)
				sFiltersAndCalcs = sFiltersAndCalcs + componentParameter(sTag, "FILTERID") + "	";
			}

			if (iType == 3) {
				// Calc (might include prompts)
				sFiltersAndCalcs = sFiltersAndCalcs + componentParameter(sTag, "CALCULATIONID") + "	";
			}

			if (iType == 1) {
				// Field (might include prompts in the child field filter)
				if (componentParameter(sTag, "FIELDSELECTIONFILTER") > 0) {
					sFiltersAndCalcs = sFiltersAndCalcs + componentParameter(sTag, "FIELDSELECTIONFILTER") + "	";
				}
			}
		}
	});
	
	var postData = {
		type: frmSend.txtSend_type.value,
		components1: frmSend.txtSend_components1.value,
		tableID: frmUseful.txtTableID.value,
		prompts: sPrompts,
		filtersAndCalcs: sFiltersAndCalcs,
		__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
	}

	$('#divValidateExpression').dialog("open");
	OpenHR.submitForm(null, "divValidateExpression", null, postData, "util_test_expression_pval");
	return true;

}

function okClick() {	
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmSend = OpenHR.getForm("divDefExpression", "frmSend");

	menu_disableMenu();

	switch (frmUseful.txtUtilType.value) {
		case "11":
			// Filter
			frmSend.txtSend_reaction.value = "FILTERS";
			break;
		case "12":
			// Calculation
			frmSend.txtSend_reaction.value = "CALCULATIONS";
			break;
		default:
			window.location.href = "defsel";
			return false;
	}

	submitDefinition();
	return true;
}

function cancelClick() {	
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	if (definitionChanged() == false) {
		menu_loadDefSelPage(frmUseful.txtUtilType.value, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, true);
	}
	else {
		OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
			if (answer == 1) {  // OK
				menu_loadDefSelPage(frmUseful.txtUtilType.value, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, true);
			}
		});
	}
	return (false);
}

function clipboardClick() {
	printClick(false);
}

function cutComponents() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	frmUseful.txtUndoType.value = "CUT";
	frmUseful.txtCutCopyType.value = "CUT";
	$('#' + tree_selectedNodeID()).css('opacity', 0.5);
	$.jstree._focused().cut();
}

function copyComponents() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	frmUseful.txtUndoType.value = "COPY";
	frmUseful.txtCutCopyType.value = "COPY";
	$.jstree._focused().copy();
}


function pasteComponents() {
	//NB Pasting is also bound to the resetIDandTag function
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	if ((frmUseful.txtCutCopyType.value != "COPY") && (frmUseful.txtCutCopyType.value != "CUT")) return true;
	
	$('#SSTree1 .jstree-leaf').css('opacity', 1);

	createUndoView("PASTE");
	
	if (tree_selectedNodeID().substr(0, 1) == "E") {
		$.jstree._focused().paste();
	} else {
		//create sibling
		$.jstree._focused().paste($('#' + tree_SelectedItemParentKey()));
	}

	refreshControls();
}

function moveComponentUp() {
	createUndoView("MOVEUP");
	var selectedID = tree_SelectedItemKey();
	var previousObj = $.jstree._focused()._get_prev('#' + selectedID, true);

	if (previousObj) {
		$("#SSTree1").jstree("move_node", "#" + selectedID, "#" + previousObj.attr('id'), "before");
		refreshControls();
	}
}

function moveComponentDown() {
	createUndoView("MOVEDOWN");
	var selectedID = tree_SelectedItemKey();
	var nextObj = $.jstree._focused()._get_next('#' + selectedID, true);

	if (nextObj) {
		$("#SSTree1").jstree("move_node", "#" + selectedID, "#" + nextObj.attr('id'), "after");
		refreshControls();
	}
}

function undoClick() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	if ((frmUseful.txtUndoType.value == "CUT") || (frmUseful.txtUndoType.value == "COPY")) {
		$('#SSTree1 .jstree-leaf').css('opacity', 1);
	} else {
		if (window.SSTree1UndoData) {
			$.jstree.rollback(window.SSTree1UndoData);
			window.SSTree1UndoData = null;
			window.SSTree1UndoData = $('#SSTree1').jstree('get_rollback');

			$('#SSTree1').jstree('destroy');
			buildjsTree();
		}
	}

	frmUseful.txtUndoType.value = "";
	frmUseful.txtCutCopyType.value = "";	//disable 'paste option'

	refreshControls();

	
}

function createUndoView(psType) {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	if ((psType != "CUT") && (psType != "COPY") && (psType != "PASTE")) frmUseful.txtCutCopyType.value = ""; //reset copy/paste values.
	frmUseful.txtUndoType.value = psType;
	window.SSTree1UndoData = $('#SSTree1').jstree('get_rollback');
	frmUseful.txtChanged.value = 1;
}

function saveChanges() {	
	cancelComponent();

	if (definitionChanged() == false) {
		$("#workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");
		return 6; // No changes made. Continue navigation
	} else {
		return 0;
	}
}

function definitionChanged() {
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
		return false;
	}

	if (frmUseful.txtChanged.value == 1) {
		return true;
	}

	return false;
}

function submitDefinition() {
	
	if (validateExpression() == false) { menu_refreshMenu(); return false; }
	if (populateSendForm() == false) { menu_refreshMenu(); return false; }

	// first populate the validate fields
	var frmSend = OpenHR.getForm("divDefExpression", "frmSend");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmOriginalDefinition = OpenHR.getForm("divDefExpression", "frmOriginalDefinition");

	var postData = {
		Action: "validate",
		validatePass: 1,
		validateName: frmDefinition.txtName.value,
		validateOwner: frmDefinition.txtOwner.value,
		validateTimestamp: ((frmUseful.txtAction.value.toUpperCase() === "EDIT") ? frmOriginalDefinition.txtDefn_Timestamp.value : 0),
		validateUtilID: ((frmUseful.txtAction.value.toUpperCase() === "EDIT") ? frmUseful.txtUtilID.value : 0),
		validateUtilType: frmSend.txtSend_type.value,
		validateAccess: frmSend.txtSend_access.value,
		components1: frmSend.txtSend_components1.value,
		validateBaseTableID: frmUseful.txtTableID.value,
		validateOriginalAccess: frmOriginalDefinition.txtOriginalAccess.value,
		__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
	}

	$('#divValidateExpression').dialog("open");
	OpenHR.submitForm(null, "divValidateExpression", null, postData, "util_validate_expression");
	return true;

}



function reEnableControls() {

	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	if (frmUseful.txtAction.value.toUpperCase() != "VIEW") {
		text_disable(frmDefinition.txtName, false);
		textarea_disable(frmDefinition.txtDescription, false);
	}

	refreshControls();

	//button_disable(frmDefinition.cmdCancel, false);
	menu_toolbarEnableItem('mnutoolCancelReport', true);
	button_disable(frmDefinition.cmdPrint, false);

	if (frmUseful.txtUtilType.value == 11) {
		button_disable(frmDefinition.cmdTest, false);
	}

	// Get menu.asp to refresh the menu.
	menu_refreshMenu();
}

function validateExpression() {
	var sTypeName;
	var sMsg;		
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	switch (frmUseful.txtUtilType.value) {
		case "11":
			// Filter
			sTypeName = "filter";
			break;
		case "12":
			// Calculation
			sTypeName = "calculation";
			break;
		default:
			sTypeName = "expression";
	}

	// Check name has been entered.
	if (frmDefinition.txtName.value == "") {		
		OpenHR.modalMessage("You must enter a name for this definition.");
		return false;
	}

	// Check the expression does have some components.      	
	if ($('#SSTree1 li').length <= 1) {
		sMsg = " The " + sTypeName + " must have some components.";
		OpenHR.modalMessage(sMsg);
		return false;
	}

	// Check that all function parameters have some components.      
	var fOK = true;

	$('#SSTree1 li[id^="E"]').each(function () {
		if ($(this).children().length == 0) {
			OpenHR.modalMessage("Function parameters must have components.");
			fOK = false;
		}
	});

	if (fOK == false) return false;

	return true;
}

function populateSendForm() {	
	var sNames = "";
	var sComponents = "";
	var reQuote = new RegExp("\"", "gi");
	
	var frmSend = OpenHR.getForm("divDefExpression", "frmSend");
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("workframe", "frmDefinition");

	// Copy all the header information to frmSend
	frmSend.txtSend_ID.value = frmUseful.txtUtilID.value;
	frmSend.txtSend_type.value = frmUseful.txtUtilType.value;
	frmSend.txtSend_name.value = frmDefinition.txtName.value;
	frmSend.txtSend_description.value = frmDefinition.txtDescription.value;
	frmSend.txtSend_userName.value = frmDefinition.txtOwner.value;

	if (frmDefinition.optAccessRW.checked == true) {
		frmSend.txtSend_access.value = "RW";
	}
	if (frmDefinition.optAccessRO.checked == true) {
		frmSend.txtSend_access.value = "RO";
	}
	if (frmDefinition.optAccessHD.checked == true) {
		frmSend.txtSend_access.value = "HD";
	}
	
	// Now go through the components	
	if ($('#SSTree1').children().length > 0) {
		var objNode = $('#SSTree1 li.root>ul>li').first();

		sComponents = "ROOT	" + objNode.attr('id') + "	" + objNode.attr('data-tag');
		sComponents = sComponents + populateSendForm_subNodes(objNode.attr('id'));

		sNames = tree_Nodetext(objNode) +
				populateSendForm_names(objNode.attr('id'));
		
		$('#SSTree1 li.root>ul>li:not(:first)').each(function () {
			sNames += "\t" + tree_Nodetext($(this));
			sComponents = sComponents + "	ROOT	" + $(this).attr('id') + "	" + $(this).attr('data-tag');
			sComponents = sComponents + populateSendForm_subNodes($(this).attr('id'));
		});

		sComponents = sComponents + "	";
	}
	
	frmSend.txtSend_components1.value = sComponents;
	frmSend.txtSend_names.value = sNames;

	frmSend.txtSend_components1.value = frmSend.txtSend_components1.value.replace(reQuote, '&quot;');
	return true;

}

function populateSendForm_subNodes(psKey) {
	var sComponents = "";
	var objNode;

	if ($('#' + psKey + '>ul>li').length > 0) {
		objNode = $('#' + psKey + '>ul>li:first');
		sComponents = "	" + psKey + "	" + objNode.attr('id') + "	" + objNode.attr('data-tag');
		sComponents = sComponents + populateSendForm_subNodes(objNode.attr('id'));


		$('#SSTree1 #' + psKey + '>ul>li:not(:first)').each(function () {
			objNode = $(this);
			sComponents += "	" + psKey + "	" + objNode.attr('id') + "	" + objNode.attr('data-tag');
			sComponents += populateSendForm_subNodes(objNode.attr('id'));
		});

	}

	return sComponents;
}

function populateSendForm_names(psKey) {
	var sNames = "";
	var objNode;	

	if ($('#' + psKey + '>ul>li').length > 0) {
		objNode = $('#' + psKey + '>ul>li').first();
		sNames = "	" + tree_Nodetext(objNode) +
				populateSendForm_names(objNode.attr('id'));

		$('#SSTree1 #' + psKey + '>ul>li:not(:first)').each(function () {
			objNode = $(this);
			sNames += "	" + tree_Nodetext(objNode) +
								populateSendForm_names(objNode.attr('id'));
		});

	}

	return sNames;
}

function ude_createNew() {
	
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	frmUseful.txtUtilID.value = 0;
	frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
	frmUseful.txtAction.value = "new";

	submitDefinition();
}

function ude_makeHidden() {

	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	frmDefinition.optAccessHD.checked = true;
	submitDefinition();
}

function SSTree1_afterLabelEdit() {

	var pfCancel = arguments[0];
	var psNewText = arguments[1];
	var sText = new String(psNewText);
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	// Remove leading spaces.
	while (sText.substr(0, 1) == " ") {
		sText = sText.substr(1);
	}

	if (sText.length == 0) {
		OpenHR.messageBox("You must enter a name.");
		pfCancel.Value = true;
		return true;
	}

	frmUseful.txtChanged.value = 1;

	refreshControls();

	return false;
}

function SSTree1_dblClick() {
	var sKey = tree_selectedNodeID();
	var frmDefinition = OpenHR.getForm("divDefExpression", "frmDefinition");

	if ((frmDefinition.cmdEdit.disabled == false) &&
			(sKey.substr(0, 1) != "E")) {
		editClick();
	}
}

function SSTree1_keyPress(sKeyPressed) {
	var sDefinition;
	var frmShortcutKeys = document.getElementById('frmShortcutKeys');
	var shortcutCollection = frmShortcutKeys.elements;	
	var sControlName;
	var sBaseName;
	var iIndex;
	var sKeys;
	var sKey;
	var sRelativeKey;
	var sShortcuts = new String(frmShortcutKeys.txtShortcutKeys.value);
	sShortcuts.toUpperCase();

	//var sKeyPressed = String.fromCharCode(piKeyAscii).toUpperCase();	
	if (sShortcuts.indexOf(sKeyPressed) >= 0) {
		for (var i = 0; i < shortcutCollection.length; i++) {
			sControlName = shortcutCollection.item(i).name;
			sBaseName = sControlName.substr(0, 16);
			if (sBaseName == "txtShortcutKeys_") {
				sKeys = shortcutCollection.item(i).value;

				if (sKeys.indexOf(sKeyPressed) >= 0) {
					iIndex = sControlName.substr(16);
					sDefinition = "0	0	" + $("#txtShortcutType_" + iIndex).val() +
							"								";

					if ($("#txtShortcutType_" + iIndex).val() == 2) {
						sDefinition = sDefinition + $("#txtShortcutID_" + iIndex).value;
					}

					sDefinition = sDefinition + "		";

					if ($("#txtShortcutType_" + iIndex).val() == 5) {
						sDefinition = sDefinition + $("#txtShortcutID_" + iIndex).val();
					}

					sDefinition = sDefinition + "																" +
						$("#txtShortcutName_" + iIndex).val() + "			";

					sKey = tree_SelectedItemKey(); // frmDefinition.SSTree1.SelectedItem.key;

					if (sKey.substr(0, 1) == "E") {
						sRelativeKey = sKey;
					}
					else {
						sRelativeKey = tree_SelectedItemParentKey();	// frmDefinition.SSTree1.SelectedItem.Parent.Key;
					}
					var shortcutParams = $('#txtShortcutParams_' + iIndex).val();
					//$("#txtShortcutParams_" + iIndex).value
					setComponent(sDefinition, "ADDEXPRCOMPONENT", sRelativeKey, shortcutParams);
					return;
				}
			}
		}
	}
}


function customMenu(node) {

	var items = {
		"create": false,
		"rename": false,
		"delete": false,
		"remove": false,
		"ccp": false,
		"1": {
			"label": "Add",
			action: function (obj) {
				tree_clickSelected(obj);
				addClick();
			}
		},
		"2": {
			"label": "Insert...",
			action: function (obj) {
				tree_clickSelected(obj);
				insertClick();
			}
		},
		"3": {
			"label": "Edit...",
			action: function (obj) {
				tree_clickSelected(obj);
				editClick();
			}
		},
		"4": {
			"label": "Delete...",
			action: function (obj) {
				tree_clickSelected(obj);
				deleteClick();
			}
		},
		"5": {
			"label": "Rename...",
			action: function (obj) {
				tree_clickSelected(obj);
				abExprMenu_Click('ID_Rename');
			}
		},
		"6": {
			"separator_before": true,
			"separator_after": true,
			"label": "View",
			"action": false,
			"submenu": {
				"expand": {
					"label": "Expand Nodes",
					action: function (obj) {
						abExprMenu_Click('ID_ExpandAll');
					}
				},
				"shrink": {
					"label": "Shrink Nodes",
					action: function (obj) {
						abExprMenu_Click('ID_ShrinkAll');
					}
				},
				"colourise": {
					"label": "Colour Nodes",
					action: function (obj) {
						abExprMenu_Click('ID_Colour');
					}
				}
			}
		},
		"7": {
			"separator_before": true,
			"separator_after": true,
			"label": "Send To",
			"action": false,
			"submenu": {
				"01": {
					"_disabled": true,
					"_class": "ui-state-disabled",
					"label": "Clipboard..."
				},
				"02": {
					"_disabled": false,
					"label": "Printer",
					action: function () {
						abExprMenu_Click('ID_OutputToPrinter');
					}
				}
			}
		},
		"8": {
			"label": "Cut",
			action: function () {
				abExprMenu_Click('ID_Cut');
			}
		},
		"9": {
			"label": "Copy",
			action: function () {
				abExprMenu_Click('ID_Copy');
			}
		},
		"10": {
			"label": "Paste",
			action: function () {
				abExprMenu_Click('ID_Paste');
			}
		},
		"11": {
			"separator_before": true,
			"label": "Move Up",
			action: function () {
				abExprMenu_Click('ID_MoveUp');
			}
		},
		"12": {
			"label": "Move Down",
			action: function () {
				abExprMenu_Click('ID_MoveDown');
			}
		},
		"13": {
			"separator_before": true,
			"label": "Undo",
			action: function () {
				abExprMenu_Click('ID_Undo');
			}
		}
	};
	
	var fRenamable;
	var sKey;
	var fModifiable;
	var sUndoText;
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	sKey = $(node).attr('id');

	$("#SSTree1").jstree("deselect_all");
	$('#SSTree1').jstree('select_node', '#' + sKey);

	fModifiable = (frmUseful.txtAction.value.toUpperCase() != "VIEW");

	// Popup menu on right button.
	//if (piButton == 2) {
	fRenamable = false;

	if (!(node.hasClass('root'))) {
		if (sKey.substr(0, 1) == "E") {
			fRenamable = fModifiable;
		}
	}

	// Enable/disable the required tools.

	if ($('#cmdAdd').prop('disabled') == true) {
		//ADD
		items["1"]["_disabled"] = true;
		items["1"]["_class"] = "ui-state-disabled";
	}

	if ($('#cmdInsert').prop('disabled') == true) {
		//Insert
		items["2"]["_disabled"] = true;
		items["2"]["_class"] = "ui-state-disabled";
	}

	if ($('#cmdEdit').prop('disabled') == true) {
		//Edit
		items["3"]["_disabled"] = true;
		items["3"]["_class"] = "ui-state-disabled";
	}

	if ($('#cmdDelete').prop('disabled') == true) {
		//Delete
		items["4"]["_disabled"] = true;
		items["4"]["_class"] = "ui-state-disabled";
	}

	items["5"]["_disabled"] = fRenamable ? false : true;
	items["5"]["_class"] = fRenamable ? "" : "ui-state-disabled";

	if (((frmUseful.txtCanCut.value != 1) || !fModifiable)) {
		//Cut
		items["8"]["_disabled"] = true;
		items["8"]["_class"] = "ui-state-disabled";
	}

	if (((frmUseful.txtCanCopy.value != 1) || !fModifiable)) {
		//Copy
		items["9"]["_disabled"] = true;
		items["9"]["_class"] = "ui-state-disabled";
	}
	//TODO: canpaste???
	if (((frmUseful.txtCanPaste.value != 1) || !fModifiable)) {
		//Paste
		items["10"]["_disabled"] = true;
		items["10"]["_class"] = "ui-state-disabled";
	}
	if (((frmUseful.txtCanMoveUp.value != 1) || !fModifiable)) {
		//Move Up
		items["11"]["_disabled"] = true;
		items["11"]["_class"] = "ui-state-disabled";
	}
	if (((frmUseful.txtCanMoveDown.value != 1) || !fModifiable)) {
		//Move Down
		items["12"]["_disabled"] = true;
		items["12"]["_class"] = "ui-state-disabled";
	}

	////For now just make the following diabled until the control is de-activeX'd
	//abExprMenu.Bands("popupSendTo").Tools("ID_OutputToPrinter").Enabled = false;
	//abExprMenu.Bands("popupSendTo").Tools("ID_OutputToClipboard").Enabled = false;

	// Set the undo text
	if (frmUseful.txtUndoType.value == "") {
		//Undo
		items["13"]["_disabled"] = true;
		items["13"]["_class"] = "ui-state-disabled";
	}

	if (frmUseful.txtUndoType.value == "ADD") {
		sUndoText = "Undo Add";
	}
	else {
		if (frmUseful.txtUndoType.value == "DELETE") {
			sUndoText = "Undo Delete";
		}
		else {
			if (frmUseful.txtUndoType.value == "PASTE") {
				sUndoText = "Undo Paste";
			}
			else {
				if (frmUseful.txtUndoType.value == "CUT") {
					sUndoText = "Undo Cut";
				}
				else {
					if (frmUseful.txtUndoType.value == "INSERT") {
						sUndoText = "Undo Insert";
					}
					else {
						if (frmUseful.txtUndoType.value == "MOVEUP") {
							sUndoText = "Undo Move Up";
						}
						else {
							if (frmUseful.txtUndoType.value == "MOVEDOWN") {
								sUndoText = "Undo Move Down";
							}
							else {
								if (frmUseful.txtUndoType.value == "EDIT") {
									sUndoText = "Undo Edit";
								}
								else {
									if (frmUseful.txtUndoType.value == "RENAME") {
										sUndoText = "Undo Rename";
									}
									else {
										sUndoText = "Undo";
									}
								}
							}
						}
					}
				}
			}
		}
	}

	items["13"]["label"] = sUndoText;

	//TODO:
	//	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
	//		abExprMenu.Bands("PopupReadOnly").TrackPopup(-1, -1);
	//	}
	//	else {
	//		abExprMenu.RecalcLayout();
	//		abExprMenu.Bands("popup1").TrackPopup(-1, -1);
	//	}
	//}

	return items;

}

function abExprMenu_Click(pTool) {	
	var sKey;
	var frmUseful = OpenHR.getForm("divDefExpression", "frmUseful");

	switch (pTool) {
		case "ID_Add":
			addClick();
			break;
		case "ID_Insert":
			insertClick();
			break;
		case "ID_Edit":
			editClick();
			break;
		case "ID_Delete":
			deleteClick();
			break;
		case "ID_Rename":
			// Only allow sub-expression labels to be edited.
			if (tree_SelectedItemLevel('#' + tree_SelectedItemKey()) > 1) {
				sKey = tree_SelectedItemKey();

				if ((sKey.substr(0, 1) == "E") &&
				(frmUseful.txtAction.value.toUpperCase() != "VIEW")) {
					createUndoView("RENAME");
					frmUseful.txtOldText.value = tree_Nodetext($('#' + sKey));
					//frmDefinition.SSTree1.StartLabelEdit();
					$("#SSTree1").jstree("rename", "#" + sKey);
				} 
			}
			break;
		case "ID_Copy":
			copyComponents();
			break;
		case "ID_Cut":
			cutComponents();
			break;
		case "ID_Paste":
			pasteComponents();
			break;
		case "ID_MoveUp":
			moveComponentUp();
			break;
		case "ID_MoveDown":
			moveComponentDown();
			break;
		case "ID_ExpandAll":
			$('#SSTree1').jstree('open_all');
			break;
		case "ID_ShrinkAll":
			$('#SSTree1').jstree('close_all');
			break;
		case "ID_Colour":
			if (frmUseful.txtExprColourMode.value == 2) {
				frmUseful.txtExprColourMode.value = 1;
			}
			else {
				frmUseful.txtExprColourMode.value = 2;
			}

			$('#SSTree1 li').not(':first').each(function () {
				var colour = frmUseful.txtExprColourMode.value == 2 ? getNodeColour(tree_SelectedItemLevel("#" + $(this).attr('id'))) : 'rgb(0,0,0)';
				$(this).find('a').css('color', colour);
			});
			break;
		case "ID_OutputToPrinter":
			printClick(true);
			break;
		case "ID_OutputToClipboard":
			clipboardClick();
			break;
		case "ID_Undo":
			undoClick();
	}
}


function tree_SelectedItemKey() {
	if (!($('#SSTree1 .jstree-clicked'))) tree_SelectRootNode();
	return $('#SSTree1 .jstree-clicked').parent().attr('id');
}

function tree_SelectedItemParentKey() {
	return $('#SSTree1 .jstree-clicked').parent().parent().parent().attr('id');
}

function tree_SelectedItemTag() {
	return $('#SSTree1 .jstree-clicked').parent().attr('data-tag');
}

function tree_SelectedItemLevel(nodeSelector) {
	//returns node level depth in tree
	//Pass in the optional nodeSelector to specify which node to start with, defaults to currently selected node.
	if (nodeSelector.length == 0) nodeSelector = '#SSTree1 .jstree-clicked';

	var count = 0;
	$(nodeSelector).parentsUntil("#SSTree1", 'ul').each(function () {
		count += 1;
	});

	return count;
}

function tree_NodesAdd(relative, relationship, key, text, tag) {
	//clean the text
	text = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
	//tag = tag.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

	switch (relationship) {
		case 3:
			//Previous. The Node is placed before the Node named in relative.
			$('<li id="' + key + '"><a style="font-weight: normal;" href="#">' + text + '</a></li>').insertBefore("#" + relative);
			$('#' + key).attr('data-tag', tag);
			break;
		case 4:
			//Child. The Node becomes a child of the Node named in relative.
			//$("#SSTree1").jstree("create","#C69315","after","No rename",false,true);
			if ($('#' + relative + ' ul').length == 0) $('#' + relative).append('<ul></ul>');
			$('#' + relative + ' ul').first().append('<li id="' + key + '"><a style="font-weight: normal;" href="#">' + text + '</a></li>');
			$('#' + key).attr('data-tag', tag);
			break;

	}

	//Set colour of node.
	$('#' + key + ' a').css('color', getNodeColour(tree_SelectedItemLevel("#" + key)));

	return $('#' + key);
}

function tree_NodesRemove(key) {
	//removes specified node from tree (and all sub nodes)
	$('#' + key).remove();

}

function tree_Refresh() {
	$('#SSTree1').jstree('refresh');
}

function tree_Nodetext(objNode) {
	//find the unrequired <ins> element and remove it, leaving just the text...
	var tmpNode = objNode;
	var x = tmpNode.find('a:first').find('ins');
	x.remove();

	return tmpNode.find('a:first').text();
}

function tree_NodeSetSelectedItem(key) {
	$('#' + key + '>a').click();
	return true;
}

function tree_GetNode1() {
	return $('#SSTree1 li.root>ul>li:first');
}

function tree_SelectRootNode() {
	$('#SSTree1').jstree('rename_node', '.root', frmDefinition.txtName.value);
	$(".root>a").click();
	return true;
}

function tree_ExpandNode(objNode) {
	$("#SSTree1").jstree("open_node", objNode);
	return true;
}

function tree_clickSelected(obj) {
	//left-clicks the item that was right-clicked.
	var elementId = obj.id;
	if (elementId == null) elementId = obj[0].id;
	if (elementId == null) elementId = $(obj).attr('id');

	if (elementId != null) {
		tree_NodeSetSelectedItem(elementId);
	}
}

function tree_getRootNodeID() {
	return $('.root').attr('id');
}

function tree_selectedNodeID() {
	return $('#SSTree1').jstree('get_selected').attr('id');
}

function tree_LastSiblingID() {
	var selectedID = $('#SSTree1').jstree('get_selected').attr('id');
	return $('#' + selectedID).parent().parent().find('li:last').attr('id');
}

function tree_FirstSiblingID() {
	var selectedID = $('#SSTree1').jstree('get_selected').attr('id');
	return $('#' + selectedID).parent().parent().find('li:first').attr('id');
}

function tree_selectedNodeChildCount() {
	return $.jstree._focused()._get_children().length;
}

//For reference:
//Select node : $('#SSTree1').jstree('select_node', '#E36896');
//Rename node : $('#SSTree1').jstree('rename_node', '#E36896', 'new text');
//select 2: $.jstree._focused().select_node("#C69307");

// $("#demo1").jstree("create", "#root", "inside", {"data": {"title": "Arithmetic", attr: {"data-tag": "nick"}}}, false, true);
