(function(window, $) {
	"use strict";

	var settingTitle = function() {
		return "setting Title";
	},
		settingHeading = function() {
			return "setting Heading";
		},
		userName = function() {
			return "username";
		},
		saveAsValues = function() {
			return "";
		},
		settingLocations = function() {
			return;
		},
		settingData = function() {
			return;
		},
		initialiseStyles = function() {
			return;
		},
		headerCols = function() {
			return;
		},
		setOptions = function (pb, var2, var3, var4, var5, var6, var7, var8, var9, var10, var11, var12) {
			return;
		},
		settingOptions = function (pb, var2, var3, var4, var5, var6, var7, var8, var9) {
			return;
		},
		getFile = function() {
			return false;
		},
		setPrinter = function() {
			return "";
		},
		resetDefaultPrinter = function() {
			return;
		},
		pivotSuppressBlanks = function() {
			return;
		},
		pivotDataFunction = function() {
			return;
		},
		addColumn = function(psString, piDecimals, pbThousandSep) {
			return;
		},
		addPage = function(psPageName, psTableName) {
			return;
		},
		arrayDim = function(poArray, piRowCount) {
			return;
		},
		addToArray = function(intCol, intRow, psValue) {
			return;
		},
		dataArray = function() {
			return;
		},
		complete = function() {
			return;
		},
		errorMessage = function() {
			return;
		};
	
	window.ClientDLL = {
		Username: userName,
		SettingTitle: settingTitle,
		SettingHeading: settingHeading,
		SaveAsValues: saveAsValues,
		SettingLocations : settingLocations,
		SettingData: settingData,
		InitialiseStyles: initialiseStyles,
		HeaderCols: headerCols,
		SetOptions: settingOptions,		
		SettingOptions: settingOptions,
		GetFile: getFile,
		SetPrinter: setPrinter,
		ResetDefaultPrinter: resetDefaultPrinter,
		PivotSuppressBlanks: pivotSuppressBlanks,
		PivotDataFunction: pivotDataFunction,
		AddColumn: addColumn,
		AddPage: addPage,
		ArrayDim: arrayDim,
		AddToArray: addToArray,
		DataArray: dataArray,
		Complete: complete ,
		ErrorMessage: errorMessage
	};

})(window, jQuery);