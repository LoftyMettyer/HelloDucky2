
// --------------------------------------------------
// Buttons
// --------------------------------------------------
function button_onMouseOver(obj) {
	return false;
	if(obj.className == 'btn btnselect')
	{
		obj.className='btn btnselect btnhov';
	}
	else
	{
		obj.className='btn btnhov';
	}
}
function button_onMouseOut(obj)
{
	return false;
	if(obj.className == 'btn btnselect btnhov')
	{
		obj.className='btn btnselect';
	}
	else
	{
		obj.className='btn';
	}
}
function button_onFocus(obj)
{
	return false;
	if(obj.className == 'btn btnhov')
	{
		obj.className='btn btnselect btnhov';
	}
	else
	{
		obj.className='btn btnselect';
	}
}
function button_onBlur(obj)
{
	return false;
	if(obj.className == 'btn btnselect btnhov')
	{
		obj.className='btn btnhov';
	}
	else
	{
		obj.className='btn';
	}
}

function button_disable(obj, pfDisable) {
	//modified for use with themeroller
	//obj.disabled = pfDisable;
	var objectID = obj.id;

	if (objectID) {
		if (pfDisable == true) {
			//obj.className='btn btndisabled';
			$("#" + objectID).addClass("ui-state-disabled btndisabled").prop("disabled", true);
		} else {
			//obj.className='btn';
			$("#" + objectID).removeClass("ui-state-disabled btndisabled").prop("disabled", false);
		}
	}
	
}
// --------------------------------------------------
// Hypertext Labels
// --------------------------------------------------
function hypertextLabel_onMouseOver(obj)
{
	if(obj.className == 'hypertext hypertextselect')
	{
		obj.className='hypertext hypertextselect hypertexthov';
	}
	else
	{
		obj.className='hypertext hypertexthov';
	}
}
function hypertextLabel_onMouseOut(obj)
{
	if(obj.className == 'hypertext hypertextselect hypertexthov')
	{
		obj.className='hypertext hypertextselect';
	}
	else
	{
		obj.className='hypertext';
	}
}
function hypertextLabel_onFocus(obj)
{
	if(obj.className == 'hypertext hypertexthov')
	{
		obj.className='hypertext hypertextselect hypertexthov';
	}
	else
	{
		obj.className='hypertext hypertextselect';
	}
}
function hypertextLabel_onBlur(obj)
{
	if(obj.className == 'hypertext hypertextselect hypertexthov')
	{
		obj.className='hypertext hypertexthov';
	}
	else
	{
		obj.className='hypertext';
	}
}

// --------------------------------------------------
// ARefs
// --------------------------------------------------
function hypertextARef_onMouseOver(obj)
{
	if(obj.className == 'hypertext hypertextselect')
	{
		obj.className='hypertext hypertextselect hypertexthov';
	}
	else
	{
		obj.className='hypertext hypertexthov';
	}
}
function hypertextARef_onMouseOut(obj)
{
	if(obj.className == 'hypertext hypertextselect hypertexthov')
	{
		obj.className='hypertext hypertextselect';
	}
	else
	{
		obj.className='hypertext';
	}
}
function hypertextARef_onFocus(obj)
{
	if(obj.className == 'hypertext hypertexthov')
	{
		obj.className='hypertext hypertextselect hypertexthov';
	}
	else
	{
		obj.className='hypertext hypertextselect';
	}
}
function hypertextARef_onBlur(obj)
{
	if(obj.className == 'hypertext hypertextselect hypertexthov')
	{
		obj.className='hypertext hypertexthov';
	}
	else
	{
		obj.className='hypertext';
	}
}

// --------------------------------------------------
// Checkbox 
// --------------------------------------------------
function checkbox_onMouseOver(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;

	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);

				if((objAssociated.className == 'checkbox checkboxdisabled') ||
					(objAssociated.className == 'checkboxdisabled'))
				{
					return;
				}

				if(objAssociated.className == 'checkbox checkboxselect')
				{
					objAssociated.className='checkbox checkboxselect checkboxhov';
				}
				else
				{
					objAssociated.className='checkbox checkboxhov';
				}

				break;
			}
		}	
	}
}
function checkbox_onMouseOut(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if((objAssociated.className == 'checkbox checkboxdisabled') ||
					(objAssociated.className == 'checkboxdisabled'))
				{
					return;
				}

				if(objAssociated.className == 'checkbox checkboxselect checkboxhov')
				{
					objAssociated.className='checkbox checkboxselect';
				}
				else
				{
					objAssociated.className='checkbox';
				}

				break;
			}
		}	
	}
}
function checkbox_onFocus(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if(objAssociated.className == 'checkbox checkboxhov')
				{
					objAssociated.className='checkbox checkboxselect checkboxhov';
				}
				else
				{
					objAssociated.className='checkbox checkboxselect';
				}

				break;
			}
		}	
	}
}
function checkbox_onBlur(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if(objAssociated.className == 'checkbox checkboxselect checkboxhov')
				{
					objAssociated.className='checkbox checkboxhov';
				}
				else
				{
					objAssociated.className='checkbox';
				}

				break;
			}
		}	
	}
}

function checkbox_disable(obj, pfDisable)
{
	if (obj.disabled == pfDisable)
	{
		return;
	}

	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;

	obj.disabled = pfDisable;

	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if (pfDisable == true)
				{
					objAssociated.className='checkbox checkboxdisabled';
				}
				else
				{
					objAssociated.className='checkbox';
				}

				break;
			}
		}	
	}
}

// --------------------------------------------------
// Checkbox Labels
// --------------------------------------------------
function checkboxLabel_onMouseOver(obj)
{
	if((obj.className == 'checkbox checkboxdisabled') ||
		(obj.className == 'checkboxdisabled'))
	{
		return;
	}

	if(obj.className == 'checkbox checkboxselect')
	{
		obj.className='checkbox checkboxselect checkboxhov';
	}
	else
	{
		obj.className='checkbox checkboxhov';
	}
}
function checkboxLabel_onMouseOut(obj)
{
	if((obj.className == 'checkbox checkboxdisabled') ||
		(obj.className == 'checkboxdisabled'))
	{
		return;
	}

	if(obj.className == 'checkbox checkboxselect checkboxhov')
	{
		obj.className='checkbox checkboxselect';
	}
	else
	{
		obj.className='checkbox';
	}
}
function checkboxLabel_onFocus(obj)
{
	if((obj.className == 'checkbox checkboxdisabled') ||
		(obj.className == 'checkboxdisabled'))
	{
		return;
	}

	if(obj.className == 'checkbox checkboxhov')
	{
		obj.className='checkbox checkboxselect checkboxhov';
	}
	else
	{
		obj.className='checkbox checkboxselect';
	}
}
function checkboxLabel_onBlur(obj)
{
	if((obj.className == 'checkbox checkboxdisabled') ||
		(obj.className == 'checkboxdisabled'))
	{
		return;
	}
	
	if(obj.className == 'checkbox checkboxselect checkboxhov')
	{
		obj.className='checkbox checkboxhov';
	}
	else
	{
		obj.className='checkbox';
	}
}
function checkboxLabel_onKeyPress(obj)
{
	try
	{
		var sCheckBoxName = obj.htmlFor;

		if (sCheckBoxName.length = 0)
		{
			// No 'for' checkbox control
			return;
		}
		
		if (document.getElementById(sCheckBoxName).disabled)
		{
			// 'for' checkbox control exists but is disabled
			return;
		}
		
		if(window.event.keyCode == 32)
		{
			// 32 = space - toggle the checkbox value
			document.getElementById(sCheckBoxName).click();	
		}
	}
	catch(e) {}
}

// --------------------------------------------------
// Radio Labels
// --------------------------------------------------
function radioLabel_onMouseOver(obj)
{
	if((obj.className == 'radio radiodisabled') ||
		(obj.className == 'radiodisabled'))
	{
		return;
	}
	
	if(obj.className == 'radio radioselect')
	{
		obj.className='radio radioselect radiohov';
	}
	else
	{
		obj.className='radio radiohov';
	}
}
function radioLabel_onMouseOut(obj)
{
	if((obj.className == 'radio radiodisabled') ||
		(obj.className == 'radiodisabled'))
	{
		return;
	}

	if(obj.className == 'radio radioselect radiohov')
	{
		obj.className='radio radioselect';
	}
	else
	{
		obj.className='radio';
	}
}

// --------------------------------------------------
// Radio 
// --------------------------------------------------
function radio_onMouseOver(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);

				if((objAssociated.className == 'radio radiodisabled') ||
					(objAssociated.className == 'radiodisabled'))
				{
					return;
				}
				
				if(objAssociated.className == 'radio radioselect')
				{
					objAssociated.className='radio radioselect radiohov';
				}
				else
				{
					objAssociated.className='radio radiohov';
				}

				break;
			}
		}	
	}
}
function radio_onMouseOut(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);

				if((objAssociated.className == 'radio radiodisabled') ||
					(objAssociated.className == 'radiodisabled'))
				{
					return;
				}

				if(objAssociated.className == 'radio radioselect radiohov')
				{
					objAssociated.className='radio radioselect';
				}
				else
				{
					objAssociated.className='radio';
				}

				break;
			}
		}	
	}
}

function radio_onFocus(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if(objAssociated.className == 'radio radiohov')
				{
					objAssociated.className='radio radioselect radiohov';
				}
				else
				{
					objAssociated.className='radio radioselect';
				}

				break;
			}
		}	
	}
}
function radio_onBlur(obj)
{
	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;
	
	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if(objAssociated.className == 'radio radioselect radiohov')
				{
					objAssociated.className='radio radiohov';
				}
				else
				{
					objAssociated.className='radio';
				}

				break;
			}
		}	
	}
}

function radio_disable(obj, pfDisable)
{
	if (obj.disabled == pfDisable)
	{
		return;
	}

	var controlCollection = document.getElementsByTagName("label");
	var objAssociated;

	obj.disabled = pfDisable;

	if (controlCollection != null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (controlCollection.item(i).htmlFor == obj.id) {
				objAssociated = controlCollection.item(i);
				
				if (pfDisable == true)
				{
					objAssociated.className='radio radiodisabled';
				}
				else
				{
					objAssociated.className='radio';
				}

				break;
			}
		}	
	}
}

// --------------------------------------------------
// Text Input
// --------------------------------------------------
function text_disable(obj, pfDisable)
{
	obj.disabled = pfDisable;
	obj.readonly = pfDisable;
	obj.locked = pfDisable;
	
	if (pfDisable == true)
	{
		obj.className='text textdisabled';
	}
	else
	{
		obj.className='text';
	}
}

// --------------------------------------------------
// TextArea
// --------------------------------------------------
function textarea_disable(obj, pfDisable)
{
	obj.disabled = pfDisable;

	if (pfDisable == true)
	{
		obj.className='textarea disabled';
	}
	else
	{
		obj.className='textarea';
	}
}

// --------------------------------------------------
// Select (combos)
// --------------------------------------------------
function combo_disable(obj, pfDisable)
{
	obj.disabled = pfDisable;

	if (pfDisable == true)
	{
		obj.className='combo combodisabled';
	}
	else
	{
		obj.className='combo';
	}
}

// --------------------------------------------------
// TreeView
// --------------------------------------------------
function treeView_disable(obj, pfDisable)
{
	// Leave the treeview enabled, just make it look disabled.
	// Handle it all in the page itself.
	if (pfDisable == true)
	{
		obj.ForeColor = 11375765;
		obj.BackColor = 15004669;
	}
	else
	{
		obj.ForeColor = 6697779;
		obj.BackColor = 15988214;
	}
}

// --------------------------------------------------
// Grid
// --------------------------------------------------
function grid_disable(obj, pfDisable)
{
	try
	{
		with (obj)
		{
			if (pfDisable == true)
			{
				HeadStyleSet("ssetFixHeaderDisabled");
				StyleSet("ssetFixDataDisabled");
				ActiveRowStyleSet("ssetSelectedDisabled");
				SelectTypeRow = 0;
			}
			else
			{
				HeadStyleSet("ssetFixHeader");
				StyleSet("ssetFixData");
				ActiveRowStyleSet("ssetSelected");
				SelectTypeRow = 1;
				RowNavigation = 1;
			}
		}
	}
	catch(e) {}
}


// --------------------------------------------------
// Image
// --------------------------------------------------
function image_disable(obj, pfDisable)
{
	obj.disabled = pfDisable;
}


// --------------------------------------------------
// Generic handler
// --------------------------------------------------
function control_disable(eElem, pfDisable) {
    
    if (eElem != null) {
        if ("text" == eElem.type) {
            text_disable(eElem, pfDisable);
        }
        else if ("TEXTAREA" == eElem.tagName) {
            textarea_disable(eElem, pfDisable);
        }
        else if ("checkbox" == eElem.type) {
            checkbox_disable(eElem, pfDisable);
        }
        else if ("radio" == eElem.type) {
            radio_disable(eElem, pfDisable);
        }
        else if ("button" == eElem.type) {
            button_disable(eElem, pfDisable);
        }
        else if ("SELECT" == eElem.tagName) {
            combo_disable(eElem, pfDisable);
        }
        else {
            grid_disable(eElem, pfDisable);
        }
    }

}



