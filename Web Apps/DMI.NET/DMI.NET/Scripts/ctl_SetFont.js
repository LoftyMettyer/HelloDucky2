﻿
function setFont(obj) {
	try {
		obj.Font.Name = 'Verdana';
		obj.Font.Bold = false;
		obj.Font.Size = 10;

		obj.Refresh();
	}
	catch (e) {
	}
}

function setTreeFont(obj) {
	try {
		obj.Font.Name = 'Verdana';
		obj.Font.Bold = false;
		obj.Font.Size = 10;

		obj.ForeColor = 6697779;
		obj.BackColor = 15988214;

		obj.Appearance = 0; // 0=flat
		obj.BorderStyle = 1; // 1=FixedSingle

		obj.Refresh();
	}
	catch (e) {
	}
}

function setMenuFont(obj) {

	try {
		obj.MenuFontStyle = 1;
		obj.Font.Name = 'Verdana';
		obj.Font.Bold = false;
		obj.Font.Size = 10;

		obj.ControlFont.Name = 'Verdana';
		obj.ControlFont.Bold = false;
		obj.ControlFont.Size = 9;

		obj.ForeColor = 6697779;
		obj.BackColor = 16248553;

		obj.Refresh();
	}
	catch (e) {
	}
}

function setGridFont(obj) {
	try {

		obj.ForeColorEven = 6697779;
		obj.ForeColorOdd = 6697779;
		obj.BackColorEven = 15988214;
		obj.BackColorOdd = 15988214;
		obj.BevelColorFrame = 10720408;
		obj.BevelColorHighlight = 16249587;
		obj.BevelColorShadow = 16249587;
		obj.BevelColorFace = 16249587;
		obj.BackColor = 16777215;

		obj.StyleSets.Add('ssetFixHeader');	// ssetHeaderEnabled
		obj.StyleSets('ssetFixHeader').ForeColor = 6697779;	// -2147483640
		obj.StyleSets('ssetFixHeader').BackColor = 16248553;	// -2147483633
		obj.HeadStyleSet('ssetFixHeader');

		obj.StyleSets.Add('ssetFixData');	// ssetEnabled
		obj.StyleSets('ssetFixData').ForeColor = 6697779;	// -2147483640
		obj.StyleSets('ssetFixData').BackColor = 15988214;	// -2147483643
		obj.StyleSet('ssetFixData');

		obj.StyleSets.Add('ssetSelected');	// ssetSelected
		obj.StyleSets('ssetSelected').ForeColor = 2774907;	// -2147483634
		obj.StyleSets('ssetSelected').BackColor = 10480637;	// -2147483635
		obj.ActiveRowStyleSet('ssetSelected');

		obj.StyleSets.Add('ssetFixHeaderDisabled');	// ssetHeaderDisabled
		obj.StyleSets('ssetFixHeaderDisabled').ForeColor = 6697779;	// -2147483631
		obj.StyleSets('ssetFixHeaderDisabled').BackColor = 16248553;	// -2147483633

		obj.StyleSets.Add('ssetFixDataDisabled'); //ssetDisabled
		obj.StyleSets('ssetFixDataDisabled').ForeColor = 11375765; //-2147483631
		obj.StyleSets('ssetFixDataDisabled').BackColor = 15004669; //-2147483633

		obj.StyleSets.Add('ssetSelectedDisabled');
		obj.StyleSets('ssetSelectedDisabled').ForeColor = 11375765;
		obj.StyleSets('ssetSelectedDisabled').BackColor = 15004669;

		for (var i = 0; i < obj.StyleSets.Count; i++) {
			obj.StyleSets(i).Font.Name = 'Verdana';
			obj.StyleSets(i).Font.Size = 10;
			obj.StyleSets(i).Font.Bold = false;
		}

		obj.RowHeight = 19;

		obj.Refresh();
	}
	catch (e) {
	}
}


