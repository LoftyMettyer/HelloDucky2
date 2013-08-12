﻿
var frmOriginalDefinition = OpenHR.getForm("workframe", "frmOriginalDefinition");
var frmDefinition = OpenHR.getForm("workframe", "frmDefinition");
var frmUseful = OpenHR.getForm("workframe", "frmUseful");

function util_def_expression_onload() {

	var fOK;
	var objNode;

	fOK = true;

	$("#workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");

	var frmUseful = OpenHR.getForm("workframe", "frmUseful");
	var frmDefinition = OpenHR.getForm("workframe", "frmDefinition");
	var sErrMsg = frmUseful.txtErrorDescription.value;
	if (sErrMsg.length > 0) {
		fOK = false;
		OpenHR.messageBox(sErrMsg);
	}

	setTreeFont(frmDefinition.SSTree1);
	setTreeFont(frmDefinition.SSTreeClipboard);
	setTreeFont(frmDefinition.SSTreeUndo);

	if (fOK == true) {
		setMenuFont(abExprMenu);

		abExprMenu.Attach();
		abExprMenu.DataPath = "include\\exprmenu.htm";
		abExprMenu.RecalcLayout();

		if (frmUseful.txtAction.value.toUpperCase() == "NEW") {
			frmUseful.txtUtilID.value = 0;
			frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
			frmDefinition.txtDescription.value = "";

			objNode = frmDefinition.SSTree1.Nodes.Add();
			sKey = "E0";
			objNode.key = sKey;
			objNode.text = "";
			objNode.tag = "";
			objNode.font.Bold = true;
			objNode.expanded = true;
		} else {
			loadDefinition();
		}
		try {
			frmDefinition.txtName.focus();
		} catch (e) {
		}

		refreshControls();
		frmUseful.txtLoading.value = 'N';
		try {
			frmDefinition.txtName.focus();
		} catch (e) {
		}

		// Get menu.asp to refresh the menu.
		menu_refreshMenu();
	}
}

function loadDefinition() {
	var sKey;

	var dataCollection = frmOriginalDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 10);
			if (sControlName == "txtDefn_E_") {
				sExprDefn = dataCollection.item(i).value;
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

					sAccess = expressionParameter(sExprDefn, "ACCESS")
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

					objNode = frmDefinition.SSTree1.Nodes.Add();
					sKey = "E" + expressionParameter(sExprDefn, "EXPRID");
					objNode.key = sKey;
					objNode.text = frmDefinition.txtName.value;
					objNode.tag = sExprDefn;
					objNode.font.Bold = true;
					objNode.expanded = true;

					// Load the expression definition into the treeview.
					loadComponentNodes(expressionParameter(sExprDefn, "EXPRID"), true);
					setInitialExpandedNodes();
					break;
				}
			}
		}
	}

	// If its read only, disable everything.
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
		disableAll();
		button_disable(frmDefinition.cmdPrint, false);
		if (frmUseful.txtUtilType.value == 11) {
			button_disable(frmDefinition.cmdTest, false);
		}
	}
}

function loadComponentNodes(piExprID, pfVisible) {
	var i;
	var sParentKey = "E" + piExprID;
	var sControlName;
	var sComponentDefn;
	var objNode;
	var fChildrenVisible;

	var dataCollection = frmOriginalDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 10);
			if (sControlName == "txtDefn_C_") {
				sComponentDefn = dataCollection.item(i).value;
				if (componentParameter(sComponentDefn, "EXPRID") == piExprID) {
					/* Load node and then load sub-expressions */
					objNode = frmDefinition.SSTree1.Nodes.Add(sParentKey, 4, "C" + componentParameter(sComponentDefn, "COMPONENTID"), componentDescription(sComponentDefn));
					//					objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sParentKey), 4, "C" + componentParameter(sComponentDefn, "COMPONENTID"), componentDescription(sComponentDefn));
					objNode.tag = sComponentDefn;

					if (pfVisible == true) {
						objNode.EnsureVisible();
					}
					objNode.foreColor = getNodeColour(objNode.level);

					if (componentParameter(sComponentDefn, "EXPANDEDNODE") == "1") {
						fChildrenVisible = true;
					}
					else {
						fChildrenVisible = false;
					}

					loadSubExpressionsNodes(componentParameter(sComponentDefn, "COMPONENTID"), fChildrenVisible);
				}
			}
		}
	}
}

function loadSubExpressionsNodes(piComponentID, pfVisible) {
	var i;
	var sControlName;
	var sExprDefn;
	var objNode;
	var fChildrenVisible;

	var sParentKey = "C" + piComponentID;

	var dataCollection = frmOriginalDefinition.elements;
	if (dataCollection != null) {
		for (i = 0; i < dataCollection.length; i++) {
			sControlName = dataCollection.item(i).name;
			sControlName = sControlName.substr(0, 10);
			if (sControlName == "txtDefn_E_") {
				sExprDefn = dataCollection.item(i).value;
				if (expressionParameter(sExprDefn, "PARENTCOMPONENTID") == piComponentID) {
					/* Load node and then load components */
					objNode = frmDefinition.SSTree1.Nodes.Add(sParentKey, 4, "E" + expressionParameter(sExprDefn, "EXPRID"), expressionParameter(sExprDefn, "NAME"));
					//					objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sParentKey), 4, "E" + expressionParameter(sExprDefn, "EXPRID"), expressionParameter(sExprDefn, "NAME"));
					objNode.tag = sExprDefn;

					if (pfVisible == true) {
						objNode.EnsureVisible();
					}
					objNode.foreColor = getNodeColour(objNode.level);

					if (expressionParameter(sExprDefn, "EXPANDEDNODE") == "1") {
						fChildrenVisible = true;
					}
					else {
						fChildrenVisible = false;
					}

					loadComponentNodes(expressionParameter(sExprDefn, "EXPRID"), fChildrenVisible);
				}
			}
		}
	}
}

function setInitialExpandedNodes() {
	var i;

	switch (frmUseful.txtExprNodeMode.value) {
		case "1":
			// Minimized.
			for (i = 1; i <= frmDefinition.SSTree1.Nodes.Count; i++) {
				if (frmDefinition.SSTree1.Nodes(i).level > 1) {
					frmDefinition.SSTree1.Nodes(i).Expanded = false;
				}
			}

			break;

		case "2":
			// Expand All.
			frmDefinition.SSTree1.focus();
			for (i = 1; i <= frmDefinition.SSTree1.Nodes.Count; i++) {
				frmDefinition.SSTree1.Nodes(i).EnsureVisible();
			}
			frmDefinition.SSTree1.Nodes(1).EnsureVisible();

			break;

		case "4":
			// Expand Top Level.
			frmDefinition.SSTree1.focus();
			for (i = 1; i <= frmDefinition.SSTree1.Nodes.Count; i++) {
				if (frmDefinition.SSTree1.Nodes(i).level <= 2) {
					frmDefinition.SSTree1.Nodes(i).EnsureVisible();
				}
				else {
					frmDefinition.SSTree1.Nodes(i).Expanded = false;
				}
			}
			frmDefinition.SSTree1.Nodes(1).EnsureVisible();

			break;
	}
}

function getNodeColour(piLevel) {
	var iColour;
	var iModLevel;

	iColour = 6697779;

	if (frmUseful.txtExprColourMode.value == 2) {
		iModLevel = piLevel % 7;

		switch (iModLevel) {
			case 0:
				iColour = 13111040;
				break;
			case 1:
				iColour = 0;
				break;
			case 2:
				iColour = 180;
				break;
			case 3:
				iColour = 32000;
				break;
			case 4:
				iColour = 8192000;
				break;
			case 5:
				iColour = 32125;
				break;
			case 6:
				iColour = 8224000;
				break;
			default:
				iColour = 8192125;
		}
	}

	return iColour;
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
																																if (psParameter == "FIELDSELECTIONORDERNAME") return

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
	var i;
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

	if (frmDefinition.SSTree1.SelectedNodes.Count == 0) {
		// Select the root node.
		frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(1);
		frmDefinition.SSTree1.SelectedItem.Expanded = true;
		frmDefinition.SSTree1.Refresh();
	}

	// Loop through each selected node
	for (i = 1; i <= frmDefinition.SSTree1.Nodes.Count; i++) {
		if (frmDefinition.SSTree1.Nodes(i).Selected == true) {
			iNodesSelected = iNodesSelected + 1;

			if (frmDefinition.SSTree1.Nodes(i).level == 1) {
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
				sKey = frmDefinition.SSTree1.Nodes(i).key;

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
					if (frmDefinition.SSTree1.Nodes(i).LastSibling.Index == frmDefinition.SSTree1.Nodes(i).Index) {
						fDisableMoveDown = true;
					}

					if (frmDefinition.SSTree1.Nodes(i).FirstSibling.Index == frmDefinition.SSTree1.Nodes(i).Index) {
						fDisableMoveUp = true;
					}
				}
			}
		}
	}

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

	if (SSTreeClipboard.Nodes.Count == 0) {
		fDisablePaste = true;
	}

	// Enable/disable controls depending on the selected component.
	button_disable(frmDefinition.cmdAdd, fDisableAdd);
	button_disable(frmDefinition.cmdInsert, fDisableInsert);
	button_disable(frmDefinition.cmdEdit, fDisableEdit);
	button_disable(frmDefinition.cmdDelete, fDisableDelete);

	button_disable(frmDefinition.cmdOK, ((frmUseful.txtChanged.value == 0) ||
			(fViewing == true)));

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
	frmDefinition.SSTree1.Nodes(1).text = frmDefinition.txtName.value;
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function changeDescription() {
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function changeAccess() {
	frmUseful.txtChanged.value = 1;
	refreshControls();
}

function addClick() {
	var fOK;
	var sKey;
	var sRelativeKey;

	var frmOptionArea = OpenHR.getForm("optionframe", "frmGotoOption");
	var frmRefresh = OpenHR.getForm("refreshframe", "frmRefresh");
	var iFunctionID = 0;
	var iParamIndex = 0;

	fOK = true;

	OpenHR.submitForm(frmRefresh);

	frmOptionArea.txtGotoOptionPage.value = "util_def_exprComponent";
	frmOptionArea.txtGotoOptionAction.value = "ADDEXPRCOMPONENT";
	frmOptionArea.txtGotoOptionTableID.value = frmUseful.txtTableID.value;
	frmOptionArea.txtGotoOptionExprID.value = frmUseful.txtUtilID.value;

	sKey = frmDefinition.SSTree1.SelectedItem.key;
	if (sKey.substr(0, 1) == "E") {
		sRelativeKey = sKey;
		nodParameter = frmDefinition.SSTree1.SelectedItem;
	}
	else {
		sRelativeKey = frmDefinition.SSTree1.SelectedItem.Parent.Key;
		nodParameter = frmDefinition.SSTree1.SelectedItem.Parent;
	}
	frmOptionArea.txtGotoOptionLinkRecordID.value = sRelativeKey;

	if ((sRelativeKey.substr(0, 1) == "E") &&
			(nodParameter.Level > 1)) {
		iType = componentParameter(nodParameter.Parent.tag, "TYPE");
		if (iType == 2) {
			// Function parameter
			iFunctionID = componentParameter(nodParameter.Parent.tag, "FUNCTIONID");

			nodTemp = nodParameter.FirstSibling;
			for (iLoop = 1; iLoop <= nodParameter.Parent.Children; iLoop++) {
				if (nodTemp.Key == nodParameter.Key) {
					iParamIndex = iLoop;
					break;
				}
				nodTemp = nodTemp.next;
			}
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
		OpenHR.submitForm(frmOptionArea);
	}
}

function insertClick() {
	var fOK;
	var frmOptionArea = OpenHR.getForm("optionframe", "frmGotoOption");
	var frmRefresh = OpenHR.getForm("refreshframe", "frmRefresh");

	var iFunctionID = 0;
	var iParamIndex = 0;

	fOK = true;
	OpenHR.submitForm(frmRefresh);

	frmOptionArea.txtGotoOptionPage.value = "util_def_exprComponent";
	frmOptionArea.txtGotoOptionAction.value = "INSERTEXPRCOMPONENT";
	frmOptionArea.txtGotoOptionTableID.value = frmUseful.txtTableID.value;
	frmOptionArea.txtGotoOptionLinkRecordID.value = frmDefinition.SSTree1.SelectedItem.key;
	frmOptionArea.txtGotoOptionExprID.value = frmUseful.txtUtilID.value;

	sKey = frmDefinition.SSTree1.SelectedItem.key;
	if (sKey.substr(0, 1) == "E") {
		sRelativeKey = sKey;
		nodParameter = frmDefinition.SSTree1.SelectedItem;
	}
	else {
		sRelativeKey = frmDefinition.SSTree1.SelectedItem.Parent.Key;
		nodParameter = frmDefinition.SSTree1.SelectedItem.Parent;
	}

	if ((sRelativeKey.substr(0, 1) == "E") &&
			(nodParameter.Level > 1)) {
		iType = componentParameter(nodParameter.Parent.tag, "TYPE");
		if (iType == 2) {
			// Function parameter
			iFunctionID = componentParameter(nodParameter.Parent.tag, "FUNCTIONID");

			nodTemp = nodParameter.FirstSibling;
			for (iLoop = 1; iLoop <= nodParameter.Parent.Children; iLoop++) {
				if (nodTemp.Key == nodParameter.Key) {
					iParamIndex = iLoop;
					break;
				}
				nodTemp = nodTemp.next;
			}
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
		OpenHR.submitForm(frmOptionArea);
	}
}

function editClick() {
	var fOK;
	var frmOptionArea = OpenHR.getForm("optionframe", "frmGotoOption");
	var frmRefresh = OpenHR.getForm("refreshframe", "frmRefresh");
	var iFunctionID = 0;
	var iParamIndex = 0;

	fOK = true;
	OpenHR.submitForm(frmRefresh);

	frmOptionArea.txtGotoOptionPage.value = "util_def_exprComponent";
	frmOptionArea.txtGotoOptionAction.value = "EDITEXPRCOMPONENT";
	frmOptionArea.txtGotoOptionTableID.value = frmUseful.txtTableID.value;
	frmOptionArea.txtGotoOptionLinkRecordID.value = frmDefinition.SSTree1.SelectedItem.key;
	frmOptionArea.txtGotoOptionExprID.value = frmUseful.txtUtilID.value;
	frmOptionArea.txtGotoOptionExtension.value = frmDefinition.SSTree1.SelectedItem.tag;

	sKey = frmDefinition.SSTree1.SelectedItem.key;
	if (sKey.substr(0, 1) == "E") {
		sRelativeKey = sKey;
		nodParameter = frmDefinition.SSTree1.SelectedItem;
	}
	else {
		sRelativeKey = frmDefinition.SSTree1.SelectedItem.Parent.Key;
		nodParameter = frmDefinition.SSTree1.SelectedItem.Parent;
	}

	if ((sRelativeKey.substr(0, 1) == "E") &&
			(nodParameter.Level > 1)) {
		iType = componentParameter(nodParameter.Parent.tag, "TYPE");
		if (iType == 2) {
			// Function parameter
			iFunctionID = componentParameter(nodParameter.Parent.tag, "FUNCTIONID");

			nodTemp = nodParameter.FirstSibling;
			for (iLoop = 1; iLoop <= nodParameter.Parent.Children; iLoop++) {
				if (nodTemp.Key == nodParameter.Key) {
					iParamIndex = iLoop;
					break;
				}
				nodTemp = nodTemp.next;
			}
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
		OpenHR.submitForm(frmOptionArea);
	}
}

function setComponent(psComponentDefn, psAction, psLinkComponentID, psFunctionParameters) {
	var iLoop;
	var iIndex;
	var fNodeExists = false;
	var objNode;
	var sNewKey;
	var sExprName;
	var sTemp;

	// Expand the work frame and hide the option frame.

	//JIRA Bug 3017, TFS Bug 104
	//The following line was here originally but was crashing IE9
	//$("#optionframe").hide();
	//This is a known IE9 bug; see for example
	//	- http://stackoverflow.com/questions/2567398/jquery-ie-crashing-on-somediv-hide
	//  - https://www.google.co.uk/#q=jquery+hide+ie9+crash
	
	//So after trying the lines below (all of them crashed IE9) I found that the fix is the (uncommented) line
	//  - $("#optionframe").css('display', 'none');
	//	- document.getElementById("#optionframe").style.display = "none";
	$("#optionframe").attr('style', 'display: none;');
	//End of bug

	$("#workframe").show();

	frmDefinition.SSTree1.style.visibility = "visible";
	frmDefinition.SSTree1.Refresh();

	for (iLoop = 1; iLoop <= frmDefinition.SSTree1.Nodes.Count; iLoop++) {
		if (frmDefinition.SSTree1.Nodes(iLoop).Key == psLinkComponentID) {
			fNodeExists = true;
			break;
		}
	}

	frmDefinition.SSTree1.focus();

	if (fNodeExists == true) {
		if (psAction == "EDITEXPRCOMPONENT") {
			createUndoView("EDIT");
			// Add the component node for the new component definition.	
			sNewKey = getUniqueNodeKey("C");
			objNode = frmDefinition.SSTree1.Nodes.Add(psLinkComponentID, 3, sNewKey, componentDescription(psComponentDefn));
			//objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(psLinkComponentID), 3, sNewKey, componentDescription(psComponentDefn));
			objNode.tag = psComponentDefn;

			objNode.EnsureVisible();
			objNode.foreColor = getNodeColour(objNode.level);

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

					objNode = frmDefinition.SSTree1.Nodes.Add(sNewKey, 4, getUniqueNodeKey("E"), sExprName);
					//					objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sNewKey), 4, getUniqueNodeKey("E"), sExprName);
					objNode.tag = "													";
					objNode.foreColor = getNodeColour(objNode.level);
					objNode.EnsureVisible();
				}
			}

			// Remove the component node for the old component definition.	
			frmDefinition.SSTree1.Nodes.remove(psLinkComponentID);

			frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(sNewKey);
			frmDefinition.SSTree1.SelectedItem.Expanded = true;
			frmDefinition.SSTree1.Refresh();
		}

		if (psAction == "ADDEXPRCOMPONENT") {
			createUndoView("ADD");
			// Add the component node for the new component definition.	
			sNewKey = getUniqueNodeKey("C");
			objNode = frmDefinition.SSTree1.Nodes.Add(psLinkComponentID, 4, sNewKey, componentDescription(psComponentDefn));
			//			objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(psLinkComponentID), 4, sNewKey, componentDescription(psComponentDefn));
			objNode.tag = psComponentDefn;

			objNode.EnsureVisible();
			objNode.foreColor = getNodeColour(objNode.level);

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

					objNode = frmDefinition.SSTree1.Nodes.Add(sNewKey, 4, getUniqueNodeKey("E"), sExprName);
					//					objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sNewKey), 4, getUniqueNodeKey("E"), sExprName);
					objNode.tag = "													";
					objNode.foreColor = getNodeColour(objNode.level);
					objNode.EnsureVisible();
				}
			}

			frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(sNewKey);
			frmDefinition.SSTree1.SelectedItem.Expanded = true;
			frmDefinition.SSTree1.Refresh();
		}

		if (psAction == "INSERTEXPRCOMPONENT") {
			createUndoView("INSERT");
			// Add the component node for the new component definition.	
			sNewKey = getUniqueNodeKey("C");
			objNode = frmDefinition.SSTree1.Nodes.Add(psLinkComponentID, 3, sNewKey, componentDescription(psComponentDefn));
			//			objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(psLinkComponentID), 3, sNewKey, componentDescription(psComponentDefn));
			objNode.tag = psComponentDefn;

			objNode.EnsureVisible();
			objNode.foreColor = getNodeColour(objNode.level);

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

					objNode = frmDefinition.SSTree1.Nodes.Add(sNewKey, 4, getUniqueNodeKey("E"), sExprName);
					//					objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sNewKey), 4, getUniqueNodeKey("E"), sExprName);
					objNode.tag = "													";
					objNode.foreColor = getNodeColour(objNode.level);
					objNode.EnsureVisible();
				}
			}

			frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(sNewKey);
			frmDefinition.SSTree1.SelectedItem.Expanded = true;
			frmDefinition.SSTree1.Refresh();
		}
	}

	try {
		frmDefinition.txtName.focus();
	}
	catch (e) { }


	refreshControls();
	refreshGrid();
}

function cancelComponent() {
	// Expand the work frame and hide the option frame.
	$("#optionframe").hide();
	$("#workframe").show();

	frmDefinition.SSTree1.style.visibility = "visible";
	frmDefinition.SSTree1.Refresh();

	//        $("#workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");
	menu_refreshMenu();

	frmDefinition.SSTree1.focus();
	refreshControls();
	refreshGrid();
}

function getUniqueNodeKey(psType) {
	var iLoop;
	var sKey;
	var sNodeKey;
	var sKeyID;
	var iKeyID;
	var iMaxKeyID = 0;

	for (iLoop = 1; iLoop <= frmDefinition.SSTree1.Nodes.Count; iLoop++) {
		sNodeKey = frmDefinition.SSTree1.Nodes(iLoop).Key;

		sKeyID = sNodeKey.substr(1);
		iKeyID = Number(sKeyID);

		if (iKeyID > iMaxKeyID) {
			iMaxKeyID = iKeyID;
		}
	}

	sKeyID = String(iMaxKeyID + 1);
	sKey = psType + sKeyID;

	return (sKey);
}

function deleteClick() {
	// Delete the selected tree nodes.
	var i;

	createUndoView("DELETE");

	for (i = frmDefinition.SSTree1.Nodes.Count; i >= 1; i--) {
		if (frmDefinition.SSTree1.Nodes(i).Selected == true) {
			frmDefinition.SSTree1.Nodes.remove(frmDefinition.SSTree1.Nodes(i).key);
		}
	}

	refreshControls();
}

function printClick(pfToPrinter) {
	var sExprType;
	//var sAccess;
	//var i;
	//var objNode;
	//var fOK = true;
	//var sClipboardText = "";
	//var sCR = String.fromCharCode(13);
	//var sLF = String.fromCharCode(10);

	//var objPrinter = window.parent.frames("menuframe").ASRIntranetPrintFunctions;	

	//if (pfToPrinter == true) {
	//    if(objPrinter.IsOK == false) {
	//        return;
	//    }
	//}

	//// OK so far.

	//if (pfToPrinter == true) {
	//    fOK = objPrinter.PrintStart(false, frmUseful.txtUserName.value);
	//}
	//else {
	//    objPrinter.ClipboardStart();
	//}

	//if (fOK == true) {	
	//    switch (frmUseful.txtUtilType.value) {
	//    case "11":
	//        // Filter
	//        sExprType = "Filter";
	//        break;
	//    case "12":
	//        // Calculation
	//        sExprType = "Runtime Calculation";
	//    default:
	//        sExprType = "Expression";
	//    }

	//    if (pfToPrinter == true) {
	//        objPrinter.PrintHeader(sExprType + " Definition : " + frmDefinition.txtName.value);
	//        objPrinter.PrintNormal("Description : " + frmDefinition.txtDescription.value);
	//        objPrinter.PrintNormal("");
	//        objPrinter.PrintNormal("Owner : " + frmDefinition.txtOwner.value);
	//    }
	//    else {
	//        sClipboardText = sClipboardText + 
	//            sExprType + " Definition : " + frmDefinition.txtName.value + sCR + sLF + 
	//            "Description : " + frmDefinition.txtDescription.value + sCR + sLF + sCR + sLF + 
	//            "Owner : " + frmDefinition.txtOwner.value + sCR + sLF;
	//    }

	//    if (frmDefinition.optAccessHD.checked == true) {
	//        sAccess = "Hidden";
	//    }
	//    else {
	//        if (frmDefinition.optAccessRO.checked == true) {
	//            sAccess = "Read only";
	//        }
	//        else {
	//            sAccess = "Read / Write";
	//        }
	//    }

	//    if (pfToPrinter == true) {
	//        objPrinter.PrintNormal("Access : " + sAccess);
	//        objPrinter.PrintNormal("");
	//        objPrinter.PrintNormal("Base Table : " + frmUseful.txtTableName.value);
	//        objPrinter.PrintNormal();
	//        objPrinter.PrintTitle("Components");
	//    }
	//    else {
	//        sClipboardText = sClipboardText + 
	//            "Access : " + sAccess + sCR + sLF + sCR + sLF +
	//            "Base Table : " + frmUseful.txtTableName.value + sCR + sLF + sCR + sLF +
	//            "Components" + sCR + sLF + sCR + sLF;

	//        objPrinter.ClipboardSetText(sClipboardText);
	//    }

	//    if (frmDefinition.SSTree1.Nodes(1).children > 0) {
	//        objNode = frmDefinition.SSTree1.Nodes(1).child;
	//        printNode(objNode, pfToPrinter);

	//        for (i=1; i< frmDefinition.SSTree1.Nodes(1).children; i++) {
	//            objNode = objNode.next;
	//            printNode(objNode, pfToPrinter);
	//        }
	//    }

	//    if (pfToPrinter == true) {
	//        objPrinter.PrintEnd();
	//        objPrinter.PrintConfirm(sExprType + " : " + frmDefinition.txtName.value, sExprType + " Definition");
	//    }
	//}
}

function printNode(pobjNode, pfToPrinter) {
	var i;
	//var objNode;
	//var objPrinter = window.parent.frames("menuframe").ASRIntranetPrintFunctions;
	//var sKey;
	//var sType;
	//var sTypeName;
	//var sClipboardText = "";
	//var sCR = String.fromCharCode(13);
	//var sLF = String.fromCharCode(10);
	//var sTAB = String.fromCharCode(9);

	//if (pfToPrinter == true) {
	//    objPrinter.PrinterBold = false;
	//    objPrinter.CurrentX = 1000 + ((pobjNode.level - 1) * 500);
	//    objPrinter.CurrentY = objPrinter.CurrentY + 100;
	//}
	//else {
	//    for (i=1; i<pobjNode.level; i++) {
	//        sClipboardText = sClipboardText + sTAB;
	//    }
	//}

	//sTypeName = "";
	//sKey = pobjNode.key;
	//sKey = sKey.substr(0,1);
	//if (sKey == "E") {
	//    sTypeName = "Parameter : ";
	//}
	//else {
	//    sType = componentParameter(pobjNode.tag, "TYPE");
	//    if (sType == 2) {
	//        // Function
	//        sTypeName = "Function : ";
	//    }
	//    if (sType == 3) {
	//        // Calculation
	//        sTypeName = "Calculation : ";
	//    }
	//    if (sType == 10) {
	//        // Filter
	//        sTypeName = "Filter : ";
	//    }
	//}

	//if (pfToPrinter == true) {
	//    objPrinter.PrintStraight(sTypeName + pobjNode.text);
	//}
	//else {
	//    sClipboardText = sClipboardText + sTypeName + pobjNode.text + sCR + sLF;
	//    objPrinter.ClipboardSetText(objPrinter.ClipboardGetText() + sClipboardText);
	//}

	//if (pobjNode.children > 0) {
	//    objNode = pobjNode.child;
	//    printNode(objNode, pfToPrinter);

	//    for (i=1; i< pobjNode.children; i++) {
	//        objNode = objNode.next;
	//        printNode(objNode, pfToPrinter);
	//    }
	//}
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

	if (validateExpression() == false) return;
	if (populateSendForm() == false) return;

	// Create a tab-delimuted string of the prompted value definitions.
	sPrompts = "";
	sFiltersAndCalcs = "";

	for (iLoop = 1; iLoop <= frmDefinition.SSTree1.Nodes.Count; iLoop++) {
		sKey = frmDefinition.SSTree1.Nodes(iLoop).key;
		sTag = frmDefinition.SSTree1.Nodes(iLoop).tag;

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
	}

	frmTest.type.value = frmSend.txtSend_type.value;
	frmTest.components1.value = frmSend.txtSend_components1.value;
	frmTest.prompts.value = sPrompts;
	frmTest.filtersAndCalcs.value = sFiltersAndCalcs;

	sURL = "util_dialog_expression" +
			"?action=test";

	openDialog(sURL, (screen.width) / 2, (screen.height) / 3);
}

function okClick() {
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
			return;
	}

	submitDefinition();
}

function cancelClick() {

	if (definitionChanged() == false) {
		menu_loadDefSelPage(frmUseful.txtUtilType.value, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, false);
		return;
	}

	answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
	if (answer == 7) {
		// No
		menu_loadDefSelPage(frmUseful.txtUtilType.value, frmUseful.txtUtilID.value, frmUseful.txtTableID.value, false);
		return (false);
	}

	if (answer == 6) {
		// Yes
		okClick();
	}
}

function clipboardClick() {
	printClick(false);
}

function cutComponents() {
	copyComponents();
	deleteClick();
	frmUseful.txtUndoType.value = "CUT"
}

function copyComponents() {
	var i;
	var objNode;
	var sNewKey;

	// Clear the current collection of copy components.
	SSTreeClipboard.Nodes.Clear();

	// Add the selected components into the collection of copy components.
	for (i = frmDefinition.SSTree1.Nodes.Count; i >= 1; i--) {
		if (frmDefinition.SSTree1.Nodes(i).Selected == true) {
			sNewKey = getUniqueClipboardNodeKey("C");

			objNode = SSTreeClipboard.Nodes.Add();
			objNode.key = sNewKey;
			objNode.text = frmDefinition.SSTree1.Nodes(i).text;
			objNode.tag = frmDefinition.SSTree1.Nodes(i).tag;

			copySubNodes(frmDefinition.SSTree1.Nodes(i).key, sNewKey);
		}
	}
}

function copySubNodes(sFromNode, sToNode) {
	// Copy all sub-nodes from one node to another.
	var i;
	var objNode;
	var objNewNode;
	var sNewKey;
	var sType;

	if (frmDefinition.SSTree1.Nodes(sFromNode).children > 0) {
		if (sFromNode.substr(0, 1) == "E") {
			sType = "C";
		}
		else {
			sType = "E";
		}
		objNode = frmDefinition.SSTree1.Nodes(sFromNode).child;

		sNewKey = getUniqueClipboardNodeKey(sType);
		objNewNode = SSTreeClipboard.Nodes.Add(sToNode, 4, sNewKey, objNode.text);
		//		objNewNode = SSTreeClipboard.Nodes.Add(SSTreeClipboard.Nodes(sToNode), 4, sNewKey, objNode.text);
		objNewNode.tag = objNode.tag;
		copySubNodes(objNode.key, sNewKey);

		for (i = 1; i < frmDefinition.SSTree1.Nodes(sFromNode).children; i++) {
			objNode = objNode.next;

			sNewKey = getUniqueClipboardNodeKey(sType);
			objNewNode = SSTreeClipboard.Nodes.Add(sToNode, 4, sNewKey, objNode.text);
			//			objNewNode = SSTreeClipboard.Nodes.Add(SSTreeClipboard.Nodes(sToNode), 4, sNewKey, objNode.text);
			objNewNode.tag = objNode.tag;
			copySubNodes(objNode.key, sNewKey);
		}
	}
}

function getUniqueClipboardNodeKey(psType) {
	var iLoop;
	var sKey;
	var sNodeKey;
	var sKeyID;
	var iKeyID;
	var iMaxKeyID = 0;

	for (iLoop = 1; iLoop <= SSTreeClipboard.Nodes.Count; iLoop++) {
		sNodeKey = SSTreeClipboard.Nodes(iLoop).Key;

		sKeyID = sNodeKey.substr(1);
		iKeyID = Number(sKeyID);

		if (iKeyID > iMaxKeyID) {
			iMaxKeyID = iKeyID;
		}
	}

	sKeyID = String(iMaxKeyID + 1);
	sKey = psType + sKeyID;

	return (sKey);
}

function pasteComponents() {
	var i;
	var objNode;
	var objCurrentNode;
	var sCurrentType;
	var sNewKey;
	var iRelation;

	createUndoView("PASTE");
	objCurrentNode = frmDefinition.SSTree1.SelectedItem;
	sCurrentType = objCurrentNode.key;
	sCurrentType = sCurrentType.substr(0, 1);

	if (sCurrentType == "E") {
		if (objCurrentNode.children == 0) {
			iRelation = 4;
		}
		else {
			objCurrentNode = objCurrentNode.child;
			iRelation = 3;
		}
	}
	else {
		iRelation = 2;
	}

	frmDefinition.SSTree1.SelectedNodes.clear();

	for (i = SSTreeClipboard.Nodes.Count; i >= 1; i--) {
		if (SSTreeClipboard.Nodes(i).level == 1) {
			sNewKey = getUniqueNodeKey("C");
			var currentNodeKey;
			currentNodeKey = objCurrentNode.Key;

			objNode = frmDefinition.SSTree1.Nodes.Add(currentNodeKey, iRelation, sNewKey, SSTreeClipboard.Nodes(i).text);
			//			objNode = frmDefinition.SSTree1.Nodes.Add(objCurrentNode, iRelation, sNewKey, SSTreeClipboard.Nodes(i).text);
			objNode.tag = SSTreeClipboard.Nodes(i).tag;
			objNode.EnsureVisible();
			objNode.foreColor = getNodeColour(objNode.level);

			pasteSubNodes(SSTreeClipboard.Nodes(i).key, sNewKey);

			//			frmDefinition.SSTree1.SelectedNodes.add(sNewKey);
			frmDefinition.SSTree1.SelectedNodes.Add(frmDefinition.SSTree1.Nodes(sNewKey));

			objCurrentNode = frmDefinition.SSTree1.Nodes(sNewKey);
			iRelation = 2;
		}
	}

	refreshControls();
}

function pasteSubNodes(sFromNode, sToNode) {
	// Copy all sub-nodes from one node to another.
	var i;
	var objNode;
	var objNewNode;
	var sNewKey;
	var sType;

	if (SSTreeClipboard.Nodes(sFromNode).children > 0) {
		if (sFromNode.substr(0, 1) == "E") {
			sType = "C";
		}
		else {
			sType = "E";
		}
		objNode = SSTreeClipboard.Nodes(sFromNode).child;

		sNewKey = getUniqueNodeKey(sType);
		objNewNode = frmDefinition.SSTree1.Nodes.Add(sToNode, 4, sNewKey, objNode.text);
		//		objNewNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sToNode), 4, sNewKey, objNode.text);
		objNewNode.tag = objNode.tag;
		objNewNode.foreColor = getNodeColour(objNewNode.level);

		pasteSubNodes(objNode.key, sNewKey);

		for (i = 1; i < SSTreeClipboard.Nodes(sFromNode).children; i++) {
			objNode = objNode.next;

			sNewKey = getUniqueNodeKey(sType);
			objNewNode = frmDefinition.SSTree1.Nodes.Add(sToNode, 4, sNewKey, objNode.text);
			//			objNewNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sToNode), 4, sNewKey, objNode.text);
			objNewNode.tag = objNode.tag;
			objNewNode.foreColor = getNodeColour(objNewNode.level);

			pasteSubNodes(objNode.key, sNewKey);
		}
	}
}

function moveComponentUp() {
	var i;
	var sKey;
	var sNewKey;

	createUndoView("MOVEUP");

	for (i = 1; i <= frmDefinition.SSTree1.Nodes.Count; i++) {
		if (frmDefinition.SSTree1.Nodes(i).Selected == true) {
			if (frmDefinition.SSTree1.Nodes(i).level != 1) {
				sKey = frmDefinition.SSTree1.Nodes(i).key;

				if (sKey.substr(0, 1) != "E") {
					if (frmDefinition.SSTree1.Nodes(i).FirstSibling.Index != frmDefinition.SSTree1.Nodes(i).Index) {
						// Move the node up on place.
						sNewKey = moveComponent(i, "UP");
						frmDefinition.SSTree1.Nodes.remove(i);
						frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(sNewKey);
						frmDefinition.SSTree1.SelectedItem.Expanded = true;
						frmDefinition.SSTree1.Refresh();
						refreshControls();
						return;
					}
				}
			}
		}
	}
}

function moveComponentDown() {
	var i;
	var sKey;
	var sNewKey;

	createUndoView("MOVEDOWN");

	for (i = 1; i <= frmDefinition.SSTree1.Nodes.Count; i++) {
		if (frmDefinition.SSTree1.Nodes(i).Selected == true) {
			if (frmDefinition.SSTree1.Nodes(i).level != 1) {
				sKey = frmDefinition.SSTree1.Nodes(i).key;

				if (sKey.substr(0, 1) != "E") {
					if (frmDefinition.SSTree1.Nodes(i).LastSibling.Index != frmDefinition.SSTree1.Nodes(i).Index) {
						// Move the node down on place.
						sNewKey = moveComponent(i, "DOWN");
						frmDefinition.SSTree1.Nodes.remove(i);
						frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(sNewKey);
						frmDefinition.SSTree1.SelectedItem.Expanded = true;
						frmDefinition.SSTree1.Refresh();
						refreshControls();
						return;
					}
				}
			}
		}
	}
}

function moveComponent(piIndex, psDirection) {
	var objNode;
	var sNewKey;
	var sRelatedKey;
	var iRelation;

	sNewKey = getUniqueNodeKey("C");

	if (psDirection == "UP") {
		sRelatedKey = frmDefinition.SSTree1.Nodes(piIndex).previous.key;
		iRelation = 3;
	}
	else {
		sRelatedKey = frmDefinition.SSTree1.Nodes(piIndex).next.key;
		iRelation = 2;
	}

	objNode = frmDefinition.SSTree1.Nodes.Add(sRelatedKey, iRelation, sNewKey, frmDefinition.SSTree1.Nodes(piIndex).text);
	//	objNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sRelatedKey), iRelation, sNewKey, frmDefinition.SSTree1.Nodes(piIndex).text);
	objNode.tag = frmDefinition.SSTree1.Nodes(piIndex).tag;
	objNode.EnsureVisible();
	objNode.foreColor = getNodeColour(objNode.level);

	moveSubNodes(frmDefinition.SSTree1.Nodes(piIndex).key, sNewKey);

	return sNewKey;
}

function moveSubNodes(sFromNode, sToNode) {
	// Copy all sub-nodes from one node to another.
	var i;
	var objNode;
	var objNewNode;
	var sNewKey;
	var sType;


	if (frmDefinition.SSTree1.Nodes(sFromNode).children > 0) {
		if (sFromNode.substr(0, 1) == "E") {
			sType = "C";
		}
		else {
			sType = "E";
		}
		objNode = frmDefinition.SSTree1.Nodes(sFromNode).child;

		sNewKey = getUniqueNodeKey(sType);
		objNewNode = frmDefinition.SSTree1.Nodes.Add(sToNode, 4, sNewKey, objNode.text);
		//		objNewNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sToNode), 4, sNewKey, objNode.text);
		objNewNode.tag = objNode.tag;
		objNewNode.EnsureVisible();
		objNewNode.foreColor = getNodeColour(objNewNode.level);
		moveSubNodes(objNode.key, sNewKey);

		for (i = 1; i < frmDefinition.SSTree1.Nodes(sFromNode).children; i++) {
			objNode = objNode.next;

			sNewKey = getUniqueNodeKey(sType);
			objNewNode = frmDefinition.SSTree1.Nodes.Add(sToNode, 4, sNewKey, objNode.text);
			//			objNewNode = frmDefinition.SSTree1.Nodes.Add(frmDefinition.SSTree1.Nodes(sToNode), 4, sNewKey, objNode.text);
			objNewNode.tag = objNode.tag;
			objNewNode.EnsureVisible();
			objNewNode.foreColor = getNodeColour(objNewNode.level);
			moveSubNodes(objNode.key, sNewKey);
		}
	}
}

function undoClick() {
	frmUseful.txtUndoType.value = "";

	var i;
	var objNode;
	var sNewKey;

	// Clear the current collection of components.
	frmDefinition.SSTree1.Nodes.Clear();

	// Add the selected components into the collection of copy components.
	objNode = frmDefinition.SSTree1.Nodes.Add();
	objNode.key = SSTreeUndo.Nodes(1).key;
	objNode.text = frmDefinition.txtName.value;
	objNode.tag = SSTreeUndo.Nodes(1).tag;
	objNode.font.Bold = true;
	objNode.expanded = true;

	createUndoSubNodes(frmDefinition.SSTree1.Nodes(1).key, true);

	refreshControls();
}

function createUndoView(psType) {
	frmUseful.txtUndoType.value = psType;

	var i;
	var objNode;
	var sNewKey;

	frmUseful.txtChanged.value = 1;

	// Clear the current collection of copy components.
	SSTreeUndo.Nodes.Clear();

	// Add the selected components into the collection of copy components.
	objNode = SSTreeUndo.Nodes.Add();
	objNode.key = frmDefinition.SSTree1.Nodes(1).key;
	objNode.text = frmDefinition.SSTree1.Nodes(1).text;
	objNode.tag = frmDefinition.SSTree1.Nodes(1).tag;

	createUndoSubNodes(frmDefinition.SSTree1.Nodes(1).key, false)
}

function createUndoSubNodes(sKey, pfExecuteUndo) {
	// Copy all sub-nodes from one node to another.
	var i;
	var objNode;
	var objNewNode;
	var objFromTree;
	var objToTree;

	if (pfExecuteUndo == true) {
		objFromTree = SSTreeUndo;
		objToTree = frmDefinition.SSTree1;
	}
	else {
		objFromTree = frmDefinition.SSTree1;
		objToTree = SSTreeUndo;
	}

	if (objFromTree.Nodes(sKey).children > 0) {
		objNode = objFromTree.Nodes(sKey).child;
		objNewNode = objToTree.Nodes.Add(sKey, 4, objNode.key, objNode.text);
		//		objNewNode = objToTree.Nodes.Add(objToTree.Nodes(sKey), 4, objNode.key, objNode.text);
		objNewNode.tag = objNode.tag;
		objNewNode.expanded = objNode.expanded;
		objNewNode.foreColor = getNodeColour(objNewNode.level);
		createUndoSubNodes(objNode.key, pfExecuteUndo);

		for (i = 1; i < objFromTree.Nodes(sKey).children; i++) {
			objNode = objNode.next;
			objNewNode = objToTree.Nodes.Add(sKey, 4, objNode.key, objNode.text);
			//			objNewNode = objToTree.Nodes.Add(objToTree.Nodes(sKey), 4, objNode.key, objNode.text);
			objNewNode.tag = objNode.tag;
			objNewNode.expanded = objNode.expanded;
			objNewNode.foreColor = getNodeColour(objNewNode.level);
			createUndoSubNodes(objNode.key, pfExecuteUndo);
		}
	}
}

function saveChanges(psAction, pfPrompt, pfTBOverride) {
	cancelComponent();

	if (definitionChanged() == false) {
		$("workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");
		return 7; //No to saving the changes, as none have been made.
	}

	answer = OpenHR.messageBox("You have changed the current definition. Save changes ?", 3);
	if (answer == 7) {
		// No
		$("workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");
		return 7;
	}

	if (answer == 6) {
		// Yes
		$("workframe").attr("data-framesource", "UTIL_DEF_EXPRESSION");
		okClick();
	}

	return 2; //Cancel.
}

function definitionChanged() {
	if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
		return false;
	}

	if (frmUseful.txtChanged.value == 1) {
		return true;
	}

	return false;
}

function submitDefinition() {
	if (validateExpression() == false) { menu_refreshMenu(); return; }
	if (populateSendForm() == false) { menu_refreshMenu(); return; }

	// first populate the validate fields
	frmValidate.validatePass.value = 1;
	frmValidate.validateName.value = frmDefinition.txtName.value;
	frmValidate.validateAccess.value = frmSend.txtSend_access.value;
	frmValidate.validateUtilType.value = frmSend.txtSend_type.value;
	frmValidate.validateAccess.value = frmSend.txtSend_access.value;
	frmValidate.validateOwner.value = frmDefinition.txtOwner.value;
	frmValidate.components1.value = frmSend.txtSend_components1.value;
	frmValidate.validateOriginalAccess.value = frmOriginalDefinition.txtOriginalAccess.value;

	if (frmUseful.txtAction.value.toUpperCase() == "EDIT") {
		frmValidate.validateTimestamp.value = frmOriginalDefinition.txtDefn_Timestamp.value;
		frmValidate.validateUtilID.value = frmUseful.txtUtilID.value;
	}
	else {
		frmValidate.validateTimestamp.value = 0;
		frmValidate.validateUtilID.value = 0;
	}

	disableButtons();

	sURL = "util_dialog_expression" + "?action=validate";

	openDialog(sURL, 500, 170);

	reEnableControls();
}

function disableButtons() {
	text_disable(frmDefinition.txtName, true);
	textarea_disable(frmDefinition.txtDescription, true);
	radio_disable(frmDefinition.optAccessHD, true);
	radio_disable(frmDefinition.optAccessRO, true);
	radio_disable(frmDefinition.optAccessRW, true);
	treeView_disable(frmDefinition.SSTree1, true);

	button_disable(frmDefinition.cmdAdd, true);
	button_disable(frmDefinition.cmdInsert, true);
	button_disable(frmDefinition.cmdEdit, true);
	button_disable(frmDefinition.cmdDelete, true);
	button_disable(frmDefinition.cmdPrint, true);

	if (frmUseful.txtUtilType.value == 11) {
		button_disable(frmDefinition.cmdTest, true);
	}

	button_disable(frmDefinition.cmdOK, true);
	button_disable(frmDefinition.cmdCancel, true);
}

function reEnableControls() {
	if (frmUseful.txtAction.value.toUpperCase() != "VIEW") {
		text_disable(frmDefinition.txtName, false);
		textarea_disable(frmDefinition.txtDescription, false);
		treeView_disable(frmDefinition.SSTree1, false);
	}

	refreshControls();

	button_disable(frmDefinition.cmdCancel, false);
	button_disable(frmDefinition.cmdPrint, false);

	if (frmUseful.txtUtilType.value == 11) {
		button_disable(frmDefinition.cmdTest, false);
	}

	// Get menu.asp to refresh the menu.
	menu_refreshMenu();
}

function refreshGrid() {
	var sSelectedNodeKey;
	var sTopNodeKey;

	// Fault 3698
	frmDefinition.SSTree1.focus();

	sSelectedNodeKey = frmDefinition.SSTree1.SelectedItem.key;
	sTopNodeKey = frmDefinition.SSTree1.TopNode.key;

	frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(1);
	frmDefinition.SSTree1.SelectedItem.EnsureVisible();

	frmDefinition.SSTree1.TopNode = frmDefinition.SSTree1.Nodes(sTopNodeKey);

	frmDefinition.SSTree1.SelectedItem = frmDefinition.SSTree1.Nodes(sSelectedNodeKey);
	frmDefinition.SSTree1.SelectedItem.EnsureVisible();
}

function validateExpression() {
	var sTypeName;
	var sMsg;
	var sKey;
	var i;

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
		OpenHR.messageBox("You must enter a name for this definition.");
		return (false);
	}

	// Check the expression does have some components.      
	if (frmDefinition.SSTree1.Nodes.Count <= 1) {
		sMsg = " The " + sTypeName + " must have some components.";
		OpenHR.messageBox(sMsg);
		return (false);
	}

	// Check that all function parameters have some components.      
	for (i = 1; i < frmDefinition.SSTree1.Nodes.Count; i++) {
		sKey = frmDefinition.SSTree1.Nodes(i).key;
		sKey = sKey.substr(0, 1);

		if (sKey == "E") {
			if (frmDefinition.SSTree1.Nodes(i).children == 0) {
				OpenHR.messageBox("Function parameters must have components.");
				return (false);
			}
		}
	}

	return (true);
}

function populateSendForm() {
	var i;
	var sNames = "";
	var sComponents = "";
	var reQuote = new RegExp("\"", "gi");

	var frmSend = OpenHR.getForm("workframe", "frmSend");
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
	if (frmDefinition.SSTree1.Nodes(1).children > 0) {
		var objNode = frmDefinition.SSTree1.Nodes(1).child;

		sComponents = "ROOT	" + objNode.key + "	" + objNode.tag;
		sComponents = sComponents + populateSendForm_subNodes(objNode.key);
		sNames = objNode.text +
				populateSendForm_names(objNode.key);

		for (i = 1; i < frmDefinition.SSTree1.Nodes(1).children; i++) {
			objNode = objNode.next;
			sNames = sNames + "	" + objNode.text +
					populateSendForm_names(objNode.key);

			sComponents = sComponents + "	ROOT	" + objNode.key + "	" + objNode.tag;
			sComponents = sComponents + populateSendForm_subNodes(objNode.key);
		}

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
	var i;

	if (frmDefinition.SSTree1.Nodes(psKey).children > 0) {
		objNode = frmDefinition.SSTree1.Nodes(psKey).child;
		sComponents = "	" + psKey + "	" + objNode.key + "	" + objNode.tag;
		sComponents = sComponents + populateSendForm_subNodes(objNode.key);

		for (i = 1; i < frmDefinition.SSTree1.Nodes(psKey).children; i++) {
			objNode = objNode.next;
			sComponents = sComponents + "	" + psKey + "	" + objNode.key + "	" + objNode.tag;
			sComponents = sComponents + populateSendForm_subNodes(objNode.key);
		}
	}

	return sComponents;
}

function populateSendForm_names(psKey) {
	var sNames = "";
	var objNode;
	var i;

	if (frmDefinition.SSTree1.Nodes(psKey).children > 0) {
		objNode = frmDefinition.SSTree1.Nodes(psKey).child;
		sNames = "	" + objNode.text +
				populateSendForm_names(objNode.key);

		for (i = 1; i < frmDefinition.SSTree1.Nodes(psKey).children; i++) {
			objNode = objNode.next;
			sNames = sNames + "	" + objNode.text +
					populateSendForm_names(objNode.key);
		}
	}

	return sNames;
}

function createNew(pPopup) {
	pPopup.close();

	frmUseful.txtUtilID.value = 0;
	frmDefinition.txtOwner.value = frmUseful.txtUserName.value;
	frmUseful.txtAction.value = "new";

	submitDefinition();
}

function removeComponents(psNodeKeys) {
	// psnodeKeys is a tab delimited string of the 
	// node keys to remove from the expression.
	var iIndex;
	var sNodeKey;

	iIndex = psNodeKeys.indexOf("	");
	while (iIndex >= 0) {
		sNodeKey = psNodeKeys.substr(0, iIndex);
		while (sNodeKey.substr(0, 1) == " ") {
			sNodeKey = sNodeKey.substr(1);
		}
		frmDefinition.SSTree1.Nodes.remove(sNodeKey);
		psNodeKeys = psNodeKeys.substr(iIndex + 1);
		iIndex = psNodeKeys.indexOf("	");
	}
	while (psNodeKeys.substr(0, 1) == " ") {
		psNodeKeys = psNodeKeys.substr(1);
	}
	frmDefinition.SSTree1.Nodes.remove(psNodeKeys);
}

function returnToDefSel() {
	window.location.href = "defsel";
}

function makeHidden(pPopup) {
	pPopup.close();
	frmDefinition.optAccessHD.checked = true;
	submitDefinition();
}

function openDialog(pDestination, pWidth, pHeight) {
	dlgwinprops = "center:yes;" +
			"dialogHeight:" + pHeight + "px;" +
			"dialogWidth:" + pWidth + "px;" +
			"help:no;" +
			"resizable:yes;" +
			"scroll:yes;" +
			"status:no;";
	window.showModalDialog(pDestination, self, dlgwinprops);
}


function util_def_expression_addhandlers() {

	OpenHR.addActiveXHandler("SSTree1", "nodeClick", SSTree1_nodeClick);
	OpenHR.addActiveXHandler("SSTree1", "beforeLabelEdit", SSTree1_beforeLabelEdit);
	OpenHR.addActiveXHandler("SSTree1", "afterLabelEdit", SSTree1_afterLabelEdit);
	OpenHR.addActiveXHandler("SSTree1", "collapse", SSTree1_collapse);
	OpenHR.addActiveXHandler("SSTree1", "dblClick", SSTree1_dblClick);
	OpenHR.addActiveXHandler("SSTree1", "keyPress", SSTree1_keyPress);
	OpenHR.addActiveXHandler("SSTree1", "keyDown", SSTree1_keyDown);
	OpenHR.addActiveXHandler("SSTree1", "mouseUp", SSTree1_mouseUp);
	OpenHR.addActiveXHandler("abExprMenu", "DataReady", abExprMenu_DataReady);
	OpenHR.addActiveXHandler("abExprMenu", "PreCustomizeMenu", abExprMenu_PreCustomizeMenu);
	OpenHR.addActiveXHandler("abExprMenu", "Click", abExprMenu_Click);
	OpenHR.addActiveXHandler("abExprMenu", "PreSysMenu", abExprMenu_PreSysMenu);
}

function SSTree1_nodeClick(pNode) {
	refreshControls();
}

function SSTree1_beforeLabelEdit(pfCancel) {
	createUndoView("RENAME");
}

function SSTree1_afterLabelEdit() {

	var pfCancel = arguments[0];
	var psNewText = arguments[1];
	var sText = new String(psNewText);

	// Remove leading spaces.
	while (sText.substr(0, 1) == " ") {
		sText = sText.substr(1);
	}

	if (sText.length == 0) {
		OpenHR.messageBox("You must enter a name.");
		pfCancel.Value = true;
		return;
	}

	frmUseful.txtChanged.value = 1;

	refreshControls();

	return false;
}

function SSTree1_collapse(pNode) {

	if (pNode.level == 1) {
		pNode.expanded = true;
	}

	frmDefinition.SSTree1.ApproximateNodeCount = frmDefinition.SSTree1.Nodes.Count;
}

function SSTree1_dblClick() {

	var sKey;

	sKey = frmDefinition.SSTree1.SelectedItem.key;

	if ((frmDefinition.cmdEdit.disabled == false) &&
			(frmDefinition.SSTree1.Nodes.Count > 1) &&
			(frmDefinition.SSTree1.Nodes.Item(1).Selected == false) &&
			(sKey.substr(0, 1) != "E")) {
		editClick();
	}
}

function SSTree1_keyPress(piKeyAscii) {

	var sDefinition;
	var shortcutCollection = frmShortcutKeys.elements;
	var sTypeControl;
	var sControlName;
	var sBaseName;
	var iIndex;
	var sKeys;
	var sKey;
	var sRelativeKey;
	var sShortcuts = new String(frmShortcutKeys.txtShortcutKeys.value);
	sShortcuts.toUpperCase();

	var sKeyPressed = String.fromCharCode(piKeyAscii).toUpperCase();

	if (sShortcuts.indexOf(sKeyPressed) >= 0) {
		for (i = 0; i < shortcutCollection.length; i++) {
			sControlName = shortcutCollection.item(i).name;
			sBaseName = sControlName.substr(0, 16);
			if (sBaseName == "txtShortcutKeys_") {
				sKeys = shortcutCollection.item(i).value;

				if (sKeys.indexOf(sKeyPressed) >= 0) {
					iIndex = sControlName.substr(16);
					sDefinition = "0	0	" + frmShortcutKeys.elements.item("txtShortcutType_" + iIndex).value +
							"								";

					if (frmShortcutKeys.all.item("txtShortcutType_" + iIndex).value == 2) {
						sDefinition = sDefinition + frmShortcutKeys.elements.item("txtShortcutID_" + iIndex).value;
					}

					sDefinition = sDefinition + "		";

					if (frmShortcutKeys.all.item("txtShortcutType_" + iIndex).value == 5) {
						sDefinition = sDefinition + frmShortcutKeys.elements.item("txtShortcutID_" + iIndex).value;
					}

					sDefinition = sDefinition + "																" +
							frmShortcutKeys.elements.item("txtShortcutName_" + iIndex).value +
							"			"

					sKey = frmDefinition.SSTree1.SelectedItem.key;

					if (sKey.substr(0, 1) == "E") {
						sRelativeKey = sKey;
					}
					else {
						sRelativeKey = frmDefinition.SSTree1.SelectedItem.Parent.Key;
					}
					setComponent(sDefinition, "ADDEXPRCOMPONENT", sRelativeKey, frmShortcutKeys.all.item("txtShortcutParams_" + iIndex).value);
					return;
				}
			}
		}
	}
}

function SSTree1_keyDown(piButton, piShift) {
	var sButton = String.fromCharCode(piButton);

	if ((piShift & 2) == 2) {
		// CTRL pressed.

		// Paste component
		if (sButton == "V") {
			frmDefinition.cmdCancel.focus();
		}

		// Copy component
		if (sButton == "C") {
			frmDefinition.cmdCancel.focus();
		}

		// Cut component
		if (sButton == "X") {
			frmDefinition.cmdCancel.focus();
		}
	}
}

function SSTree1_mouseUp(piButton, piShift, psngX, psngY) {

	var fRenamable;
	var sKey;
	var fModifiable;
	var sUndoText;

	sKey = frmDefinition.SSTree1.SelectedItem.key;

	fModifiable = (frmUseful.txtAction.value.toUpperCase() != "VIEW");

	// Popup menu on right button.
	if (piButton == 2) {
		fRenamable = false;

		if (frmDefinition.SSTree1.SelectedItem.level > 1) {
			if (sKey.substr(0, 1) == "E") {
				fRenamable = fModifiable;
			}
		}

		// Enable/disable the required tools.
		abExprMenu.Bands("popup1").Tools("ID_Add").Enabled = (frmDefinition.cmdAdd.disabled == false);
		abExprMenu.Bands("popup1").Tools("ID_Insert").Enabled = (frmDefinition.cmdInsert.disabled == false);
		abExprMenu.Bands("popup1").Tools("ID_Edit").Enabled = (frmDefinition.cmdEdit.disabled == false);
		abExprMenu.Bands("popup1").Tools("ID_Delete").Enabled = (frmDefinition.cmdDelete.disabled == false);
		abExprMenu.Bands("popup1").Tools("ID_Rename").Enabled = fRenamable;
		abExprMenu.Bands("popup1").Tools("ID_Cut").Enabled = ((frmUseful.txtCanCut.value == 1) && fModifiable);
		abExprMenu.Bands("popup1").Tools("ID_Copy").Enabled = ((frmUseful.txtCanCopy.value == 1) && fModifiable);
		abExprMenu.Bands("popup1").Tools("ID_Paste").Enabled = ((frmUseful.txtCanPaste.value == 1) && fModifiable);
		abExprMenu.Bands("popup1").Tools("ID_MoveUp").Enabled = ((frmUseful.txtCanMoveUp.value == 1) && fModifiable);
		abExprMenu.Bands("popup1").Tools("ID_MoveDown").Enabled = ((frmUseful.txtCanMoveDown.value == 1) && fModifiable);
		abExprMenu.Bands("popup1").Tools("ID_Undo").Enabled = (frmUseful.txtUndoType.value != "");

		// Set the undo text
		abExprMenu.Bands("popup1").Tools("ID_Undo").Enabled = (frmUseful.txtUndoType.value != "");

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

		abExprMenu.Bands("popup1").Tools("ID_Undo").Caption = sUndoText;

		if (frmUseful.txtAction.value.toUpperCase() == "VIEW") {
			abExprMenu.Bands("PopupReadOnly").TrackPopup(-1, -1);
		}
		else {
			abExprMenu.RecalcLayout();
			abExprMenu.Bands("popup1").TrackPopup(-1, -1);
		}
	}
}

function abExprMenu_DataReady() {
	var sKey;
	sKey = new String("tempmenufilepath_");

	var frmMenuInfo = OpenHR.getForm("menuFrame", "frmMenuInfo");

	sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
	//  sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);

	if (sPath == "") {
		sPath = "c:\\";
	}

	if (sPath == "<NONE>") {
		frmUseful.txtMenuSaved.value = 1;
		abExprMenu.RecalcLayout();
	}
	else {
		if (sPath.substr(sPath.length - 1, 1) != "\\") {
			sPath = sPath.concat("\\");
		}

		sPath = sPath.concat("tempexpr.asp");
		if ((abExprMenu.Bands.Count() > 0) && (frmUseful.txtMenuSaved.value == 0)) {
			try {
				abExprMenu.save(sPath, "");
			}
			catch (e) {
				OpenHR.messageBox("The specified temporary menu file path cannot be written to. The temporary menu file path will be cleared.");
				sKey = new String("tempMenuFilePath_");
				sKey = sKey.concat(frmMenuInfo.txtDatabase.value);
				//    window.parent.frames("menuframe").ASRIntranetFunctions.SaveRegistrySetting("HR Pro", "DataPaths", sKey, "<NONE>");
			}

			frmUseful.txtMenuSaved.value = 1;
		}
		else {
			try {
				if ((abExprMenu.Bands.Count() == 0) && (frmUseful.txtMenuSaved.value == 1)) {
					abExprMenu.DataPath = sPath;
					abExprMenu.RecalcLayout();
					return;
				}
			}
			catch (e) { }
		}
	}
}

function abExprMenu_PreCustomizeMenu(pfCancel) {
	pfCancel = true;
	OpenHR.messageBox("The menu cannot be customized. Errors will occur if you attempt to customize it. Click anywhere in your browser to remove the dummy customisation menu.");
}

function abExprMenu_Click(pTool) {

	var iCount;
	var sKey;

	switch (pTool.name) {
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
			if (frmDefinition.SSTree1.SelectedItem.level > 1) {
				sKey = frmDefinition.SSTree1.SelectedItem.key;

				if ((sKey.substr(0, 1) == "E") &&
						(frmUseful.txtAction.value.toUpperCase() != "VIEW")) {
					frmUseful.txtOldText.value = frmDefinition.SSTree1.SelectedItem.text;
					frmDefinition.SSTree1.StartLabelEdit();
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
			for (iCount = 1; iCount <= frmDefinition.SSTree1.Nodes.Count; iCount++) {
				frmDefinition.SSTree1.Nodes(iCount).EnsureVisible();
			}
			frmDefinition.SSTree1.SelectedItem.EnsureVisible();
			break;
		case "ID_ShrinkAll":
			for (iCount = 1; iCount <= frmDefinition.SSTree1.Nodes.Count; iCount++) {
				if (frmDefinition.SSTree1.Nodes(iCount).level > 1) {
					frmDefinition.SSTree1.Nodes(iCount).Expanded = false;
				}
			}
			break;
		case "ID_ZoomIn":
			frmDefinition.SSTree1.Font.Size = frmDefinition.SSTree1.Font.Size + 2;
			for (iCount = 1; iCount <= frmDefinition.SSTree1.Nodes.Count; iCount++) {
				frmDefinition.SSTree1.Nodes(iCount).Font.Size = frmDefinition.SSTree1.Font.Size;
			}
			frmDefinition.SSTree1.SelectedItem.EnsureVisible();
			abExprMenu.Tools("ID_ZoomIn").Enabled = (frmDefinition.SSTree1.Font.Size < 11);
			abExprMenu.Tools("ID_ZoomOut").Enabled = true;
			break;
		case "ID_ZoomOut":
			frmDefinition.SSTree1.Font.Size = frmDefinition.SSTree1.Font.Size - 2;
			for (iCount = 1; iCount <= frmDefinition.SSTree1.Nodes.Count; iCount++) {
				frmDefinition.SSTree1.Nodes(iCount).Font.Size = frmDefinition.SSTree1.Font.Size;
			}
			frmDefinition.SSTree1.SelectedItem.EnsureVisible();
			abExprMenu.Tools("ID_ZoomOut").Enabled = (frmDefinition.SSTree1.Font.Size > 7);
			abExprMenu.Tools("ID_ZoomIn").Enabled = true;
			break;
		case "ID_ZoomNormal":
			frmDefinition.SSTree1.Font.Size = 8;
			for (iCount = 1; iCount <= frmDefinition.SSTree1.Nodes.Count; iCount++) {
				frmDefinition.SSTree1.Nodes(iCount).Font.Size = frmDefinition.SSTree1.Font.Size;
			}
			frmDefinition.SSTree1.SelectedItem.EnsureVisible();
			abExprMenu.Tools("ID_ZoomOut").Enabled = true;
			abExprMenu.Tools("ID_ZoomIn").Enabled = true;
			break;
		case "ID_Colour":
			if (frmUseful.txtExprColourMode.value == 2) {
				frmUseful.txtExprColourMode.value = 1;
			}
			else {
				frmUseful.txtExprColourMode.value = 2;
			}

			pTool.Checked = (frmUseful.txtExprColourMode.value == 2);
			for (iCount = 1; iCount <= frmDefinition.SSTree1.Nodes.Count; iCount++) {
				frmDefinition.SSTree1.Nodes(iCount).foreColor = getNodeColour(frmDefinition.SSTree1.Nodes(iCount).level);
			}
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

function abExprMenu_PreSysMenu(pBand) {
	if (pBand.Name == "SysCustomize") {
		pBand.Tools.RemoveAll();
	}
}


