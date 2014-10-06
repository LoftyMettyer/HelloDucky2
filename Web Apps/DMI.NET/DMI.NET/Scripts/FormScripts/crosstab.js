
function AddToIntTypeCombo(strText, strValue) {
    var oOption = document.createElement("OPTION");
    var cboIntersectionType = document.getElementById("cboIntersectionType");
    cboIntersectionType.options.add(oOption);
    oOption.innerText = strText;
    oOption.Value = strValue;
}

function AddToPgbCombo(strText, strValue) {
    var oOption = document.createElement("OPTION");
    var cboPage = document.getElementById("cboPage");
    cboPage.options.add(oOption);
    oOption.innerText = strText;
    oOption.Value = strValue;
}

function refreshCombo(psComboKey) {
    try {
        var iSelectedIndex;
        var objCombo;

        if (psComboKey == "PAGE") {
            objCombo = document.getElementById("cboPage");
        }
        else {
            if (psComboKey == "INTERSECTIONTYPE") {
                objCombo = document.getElementById("cboIntersectionType");
            }
            else {
                objCombo = document.getElementById("cboFileFormat");
            }
        }

        iSelectedIndex = objCombo.selectedIndex;

        while (cboDummy.options.length > 0) {
            cboDummy.options.remove(0);
        }

        while (objCombo.options.length > 0) {
            var oOption = document.createElement("OPTION");
            cboDummy.options.add(oOption);
            oOption.innerText = objCombo.item(0).innerText;

            // Needs both the value and the Value properties - Capital V for dropdowns to work, lowercase for drilldown to work.
            oOption.Value = objCombo.item(0).Value;
            oOption.value = objCombo.item(0).value;
            objCombo.options.remove(0);
        }

        while (cboDummy.options.length > 0) {
            var oOption = document.createElement("OPTION");
            objCombo.options.add(oOption);
            oOption.innerText = cboDummy.item(0).innerText;
            oOption.Value = cboDummy.item(0).Value;
            oOption.value = cboDummy.item(0).value;
            cboDummy.options.remove(0);
        }

        objCombo.selectedIndex = iSelectedIndex;
    }
    catch (e) { }
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
        lngPageNumber = window.cboPage.options[window.cboPage.selectedIndex].Value;
    }

    var lngIntType = 0;
    if (window.cboIntersectionType.selectedIndex != -1) {
        lngIntType = window.cboIntersectionType.options[window.cboIntersectionType.selectedIndex].Value;
    }

    var blnShowPer = (window.chkPercentType.checked == true);
    var blnPerPage = (window.chkPercentPage.checked == true);
    var blnSupZeros = (window.chkSuppressZeros.checked == true);
    var blnThousand = (window.chkUse1000.checked == true);

    getCrossTabData(strMode, lngPageNumber, lngIntType, blnShowPer, blnPerPage, blnSupZeros, blnThousand);
}
