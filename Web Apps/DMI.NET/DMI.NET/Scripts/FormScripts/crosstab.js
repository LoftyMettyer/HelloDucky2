
function AddToIntTypeCombo(strText, strValue) {
	$("#cboIntersectionType").append('<option value=' + strValue + '>' + strText + '</option>');
}

function AddToPgbCombo(strText, strValue) {
	$("#cboPage").append('<option value=' + strValue + '>' + strText + '</option>');   
}



function chkPercentType_Click() {

		checkbox_disable(chkPercentPage, (chkPercentType.checked == false));
		if (chkPercentType.checked == false) {
				chkPercentPage.checked = false;
		}
		UpdateGrid();
}

function UpdateGrid() {

		var strMode = "REFRESH";

		var lngPageNumber = 0;
		if (window.cboPage.selectedIndex != -1) {
				lngPageNumber = window.cboPage.options[window.cboPage.selectedIndex].value;
		}

		var lngIntType = 0;
		if (window.cboIntersectionType.selectedIndex != -1) {
				lngIntType = window.cboIntersectionType.options[window.cboIntersectionType.selectedIndex].value;
		}

		var blnShowPer = (window.chkPercentType.checked == true);
		var blnPerPage = (window.chkPercentPage.checked == true);
		var blnSupZeros = (window.chkSuppressZeros.checked == true);
		var blnThousand = (window.chkUse1000.checked == true);

		getCrossTabData(strMode, lngPageNumber, lngIntType, blnShowPer, blnPerPage, blnSupZeros, blnThousand);
}
