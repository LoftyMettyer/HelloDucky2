var isPermittedKeystroke = true;

function GetMaxLength(targetField) {
  // MaxLength of the control will be cover any digits, the decimal point and the minus sign.
  // Need to return the MaxLength ignoring the minus sign (ie. MaxLength -1)
  return targetField.maxLength - 1;
}

//
// Limit the text input in the specified field.
//
function WebNumericEditValidation_KeyDown(targetField, keyCode, sourceEvent) {
  // Allow non-printing, arrow and delete keys
  isPermittedKeystroke = ((keyCode < 32)
                      || (keyCode >= 33 && keyCode <= 40));

  if ((keyCode == 8) || (keyCode == 46))
    WebNumericEditValidation_KeyPress(targetField, keyCode, sourceEvent);
}

function WebNumericEditValidation_KeyPress(targetField, keyCode, sourceEvent) {
  var inputAllowed = true;
  var isPermittedValue = true;
  var enteredKeystroke;
  var maximumFieldLength;
  var currentFieldLength;
  var newValue;

  var minValue = targetField.min;
  var maxValue = targetField.max;

  var selectionLength = parseFloat(targetField.getSelectedText().length);

  if (GetMaxLength(targetField) != null) {
    // Get the current and maximum field length
    currentFieldLength = parseFloat(targetField.text.replace(targetField.minus, '').length);
    maximumFieldLength = parseFloat(GetMaxLength(targetField));

    // Allow non-printing, arrow and delete keys
    enteredKeystroke = window.event ? sourceEvent.event.keyCode : sourceEvent.event.which;

    // correct for NumPad digits
    if (enteredKeystroke >= 96 && enteredKeystroke <= 105) {
      enteredKeystroke = enteredKeystroke - 48;
    }

    var isValidKey = (isPermittedKeystroke
                    || IsMinusSign(targetField.ID, enteredKeystroke)
                    || IsDecimalSeparator(targetField.ID, enteredKeystroke));

    var isDeleteKey = (((enteredKeystroke == 8)
                    || (enteredKeystroke == 46))
                    && (sourceEvent.event.type == 'keydown'));

    if (isDeleteKey) {
      if (targetField.getSelection() > 0) {
        if (selectionLength == 0) {
          if (enteredKeystroke == 8) {
            var position = targetField.getSelection(true);
            newValue = targetField.text.substring(0, position - (selectionLength + 1)) + targetField.text.substring(position - selectionLength);
          } else {
            var position = targetField.getSelection(false);
            newValue = targetField.text.substring(0, position - selectionLength) + targetField.text.substring(position + (selectionLength + 1));
          }
        }
        else {
          if (enteredKeystroke == 8) {
            var position = targetField.getSelection(true);
            newValue = targetField.text.substring(0, position) + targetField.text.substring(position + selectionLength);
          } else {
            var position = targetField.getSelection(false);
            newValue = targetField.text.substring(0, position - selectionLength) + targetField.text.substring(position);
          }
        }
      } else
        newValue = targetField.text.substring(1);
    } else {
      var position = targetField.getSelection();
      var charValue = Chr(enteredKeystroke);
      var markerChar = '&&';
      newValue = targetField.text.substring(0, position - selectionLength) + markerChar + targetField.text.substring(position);

      if (IsDecimalSeparator(targetField.ID, enteredKeystroke))
        newValue = newValue.replace(targetField.decimalSeparator, '');

      if (IsMinusSign(targetField.ID, enteredKeystroke))
        newValue = newValue.replace(targetField.minus, '');

      newValue = newValue.replace(markerChar, charValue);
    }

    // Decide whether the keystroke is allowed to proceed
    if (!isValidKey) {
      if ((currentFieldLength - selectionLength) >= maximumFieldLength) {
        inputAllowed = false;
      }
    }

    isPermittedValue = ((newValue <= maxValue)
                         && (newValue >= minValue));

    if (!isPermittedValue) {
      if (parseFloat(newValue) > maxValue) {
        inputAllowed = false;
      }
      else if (parseFloat(newValue) < minValue) {
        inputAllowed = false;
      }
    }

    // Force a trim of the textarea contents if necessary
    if (currentFieldLength > maximumFieldLength) {
      targetField.setValue(targetField.text.substring(0, maximumFieldLength))
    }
  }

  sourceEvent.cancel = (!inputAllowed);
  return (inputAllowed);
}

function WebNumericEditValidation_Paste(targetField, sourceEvent, sID) {
  var igControl = igedit_getById(sID);
  var maxValue = igControl.max;
  var minValue = igControl.min;
  var minus = igControl.minus;

  var clipboardText = window.clipboardData.getData("Text");
  var currentFieldLength = parseFloat(targetField.value.replace(minus, '').length);
  var selectionLength = parseFloat(GetSelectionLength(targetField));
  var resultantLength = currentFieldLength + clipboardText.replace(minus, '').length - selectionLength;

  if (resultantLength > GetMaxLength(targetField))
    inputAllowed = false;

  isPermittedValue = ((parseFloat(clipboardText) <= maxValue)
                         && (parseFloat(clipboardText) >= minValue));

  if (!isPermittedValue) {
    if (parseFloat(clipboardText) >= maxValue)
      targetField.value = maxValue;
    else if (parseFloat(clipboardText) <= minValue)
      targetField.value = minValue;
    inputAllowed = false;
  }
}

//
// Returns the number of selected characters in 
// the specified element
//
function GetSelectionLength(targetField) {
  if (targetField.selectionStart == undefined) {
    return document.selection.createRange().text.length;
  }
  else {
    return (targetField.selectionEnd - targetField.selectionStart);
  }

  sourceEvent.returnValue = inputAllowed;
  return (inputAllowed);
}

function IsMinusSign(targetFieldId, checkChar) {
  var minus = igedit_getById(targetFieldId).minus;

  if (Chr(checkChar) == minus)
    return true;
  else return false;
}

function IsDecimalSeparator(targetFieldId, checkChar) {
  var decimalSeparator = igedit_getById(targetFieldId).decimalSeparator;

  if (Chr(checkChar) == decimalSeparator)
    return true;
  else return false;

}

function Asc(String) {
  return String.charCodeAt(0);
}

function Chr(AsciiNum) {
  return String.fromCharCode(AsciiNum);
}

