@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Imports System.Data
@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of OrganisationReportModel)
@Html.HiddenFor(Function(m) m.FilterColumnsAsString, New With {.id = "txtFilterColumns"})

<div class="width100" id="orgFilterSelect">
   <fieldset id="OrgFilterSelection" class="floatleft width90">
      <legend class="fontsmalltitle">Define Filter :</legend>
      <div class="nowrap tablelayout" style="margin-top: 0;">                 
            <table id="DBGridFilterRecords"></table>        
      </div>
      <div class="tablecell floatright" id="gridButton">
         <input id="btnRemove" name="btnRemove" type="button" value="Remove" style="width: 100px; margin-top: 10px;" class="btn" onclick="FilterSelect_remove()" />
         <input id="btnRemoveAll" name="btnRemoveAll" type="button" value="Remove All" style="margin-top: 10px;" class="btn" onclick="FilterSelect_removeAll()" />
      </div>
   </fieldset>   
  
   <fieldset id="OrgFilterCriteria" class="floatleft width90">
      <legend class="fontsmalltitle">Define more criteria :</legend>

      <table class="width100">
         <thead class="fontsmalltitle">
            <tr>
               <td style="width:1%"></td>
               <td style="width:30%;text-align:left">Field :</td>
               <td style="width:1%"></td>
               <td style="width:20%;text-align:left">Operator :</td>
               <td style="width:1%"></td>
               <td style="width:35%;text-align:left">Value :</td>
               <td style="width:1%"></td>
               <td ></td>               
            </tr>
         </thead>

         <tr style="">
            <td></td>
            <td>               
               <select id="selectColumn" name="selectColumn" class="combo" style="width: 100%" onchange="refreshSelectColumnCombo()"></select>
            </td>
            <td></td>
            <td>
               <input type="text" id="txtConditionLogic" Class="text textdisabled" name="selectConditionLogic" disabled="disabled" value="is equal to">

               <select id="selectConditionDate" name="selectConditionDate" Class="combo" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
                  <option value="1">is equal to</option>
                  <option value="2">is NOT equal to</option>
                  <option value="5">after</option>
                  <option value="6">before</option>
                  <option value="4">is equal to or after</option>
                  <option value="3">is equal to or before</option>
               </select>

               <select id="selectConditionNum" name="selectConditionNum" Class="combo" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
                  <option value="1">is equal to</option>
                  <option value="2">is NOT equal to</option>
                  <option value="5">is greater than</option>
                  <option value="4">is greater than or equal to</option>
                  <option value="6">is less than</option>
                  <option value="3">is less than or equal to</option>
               </select>

               <select id="selectConditionChar" name="selectConditionChar" Class="combo" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
                  <option value="1">is equal to</option>
                  <option value="2">is Not equal to</option>
                  <option value="7"> contains</option>
                  <option value="8"> does Not contain</option>
               </select>
            </td>
            <td></td>
            <td style="text-align: left">
               <select id="selectValue" name='selectValue"' class="combo" style="width: 30%;">
                  <option value="1">True</option>
                  <option value="0">False</option>
               </select>
               <input id="txtValue" name="txtValue" Class="text" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
               <input id="selectDate" name="selectDate" type="text" Class="datepicker" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;" />
            </td>
            <td></td>
            <td style="width: 175px; text-align: right">
               <input id="btnAddToList" name="btnAddToList" type="button" value="Add To List" onclick="AddToList()" />              
            </td>
         </tr>
      </table>
   </fieldset>   
   
</div>
 
<script type="text/javascript">
   $(function () {
      var colMode = [];
      var colData = [];
      var colNames = [];

      //Create the column model
      colMode.push({ name: 'ID', hidden: true });
      colMode.push({ name: 'FieldName', width: 100 });
      colMode.push({ name: 'OperatorName', width: 100 });
      colMode.push({ name: 'FilterValue', width: 100 });
      colMode.push({ name: 'FieldID', hidden: true });
      colMode.push({ name: 'OperatorID', hidden: true });
      colMode.push({ name: 'FieldDataType', hidden: true });

      colNames.push('ID');
      colNames.push('Field');
      colNames.push('Operator');
      colNames.push('Value');
      colNames.push('FieldID');
      colNames.push('OperatorID');
      colNames.push('FieldDataType');

      $('table').attr('border', '0'); //Change 0 to 1 to show borders.

      $(".datepicker").datepicker();
      $(document).on('keydown', '.datepicker', function (event) {

         switch (event.keyCode) {
            case 113:
               $(this).datepicker("setDate", new Date());
               $(this).datepicker('widget').hide('true');
               break;
         }
      });

      $(document).on('blur', '.datepicker', function (sender) {
         if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
            OpenHR.modalMessage("Invalid date value entered");
            sender.target.value = "";
            $(sender.target.id).focus();
         }
      });

      //Populate Filter Field dropdown
      GetFilterColumns();
      //$("select#selectColumn").focus();

      //Determine if the grid already exists...
      if ($("#DBGridFilterRecords").getGridParam("reccount") == undefined) { //It doesn't exist, create it
         $("#DBGridFilterRecords").jqGrid({
            multiselect: false,
            data: colData,
            datatype: 'local',
            colNames: colNames,
            colModel: colMode,
            rowNum: 1000,
            autowidth: true,
            shrinkToFit: true,
            onSelectRow: function () {
               button_disable($("#btnRemoveAll")[0], false);
               button_disable($("#btnRemove")[0], false);
            },
            editurl: 'clientArray'
         }).jqGrid('hideCol', 'cb');
      }

      if ( $("#ActionType").val() === '@UtilityActionType.Edit' || $("#ActionType").val() == '@UtilityActionType.Copy' ) {
         var datastr = '@Model.FiltersFieldList.ToJsonResult';
         GetExistingFilterDefinition(datastr);
      }
      FilterSelect_refreshControls();
   })


   function GetFilterColumns() {
      // Gets view id of selected base view
      var viewID = $("#BaseViewId").val();
      var optionOfAllType;

      $.ajax({
         url: 'Reports/GetFilterColumns?ViewID=' + viewID,
         datatype: 'json',
         mtype: 'GET',
         cache: false,
         success: function (json) {
            for (var i = 0; i < json.length; i++) {
               optionOfAllType += "<option value='" + json[i].FieldID + "' data-datatype='" + json[i].FieldDataType + "' data-size='" + json[i].FieldColumnSize + "' data-decimals='" + json[i].FieldDecimals + "'>" + json[i].FieldName + "</option>";
            }
            $("select#selectColumn").html(optionOfAllType);
         }
      });
   
   }

   function refreshSelectColumnCombo() {      
      if ($("#selectColumn option:selected")[0] != undefined) {
         var dropDown = $("#selectColumn option:selected")[0];
         var iDataType = dropDown.attributes["data-datatype"].value;
         displayOperatorSelector(iDataType);
         button_disable($("#btnAddToList")[0], isDefinitionReadOnly());
      }
   }

   function displayOperatorSelector(piDataType) {

      if (piDataType == -7) {

         $("#txtConditionLogic").css("width", "100%");
         $("#txtConditionLogic").css("visibility", "");
         $("#txtConditionLogic").css("position", "");
         $("#txtConditionLogic").css("top", "");
         $("#txtConditionLogic").css("left", "");

         $("#selectValue").css("width", "25%");
         $("#selectValue").css("visibility", "");
         $("#selectValue").css("position", "");
         $("#selectValue").css("top", "");
         $("#selectValue").css("left", "");

         $("#selectDate").css("width", "0px");
         $("#selectDate").css("visibility", "hidden");
         $("#selectDate").css("position", "absolute");
         $("#selectDate").css("top", "0");
         $("#selectDate").css("left", "0");

         $("#txtValue").css("width", "0px");
         $("#txtValue").css("visibility", "hidden");
         $("#txtValue").css("position", "absolute");
         $("#txtValue").css("top", "0");
         $("#txtValue").css("left", "0");

      } else if (piDataType == 11) {

         $("#txtConditionLogic").css("width", "0px");
         $("#txtConditionLogic").css("visibility", "hidden");
         $("#txtConditionLogic").css("position", "absolute");
         $("#txtConditionLogic").css("top", "0px");
         $("#txtConditionLogic").css("left", "0px");

         $("#selectValue").css("width", "0px");
         $("#selectValue").css("visibility", "hidden");
         $("#selectValue").css("position", "absolute");
         $("#selectValue").css("top", "0px");
         $("#selectValue").css("left", "0px");

         $("#selectDate").css("width", "100%");
         $("#selectDate").css("visibility", "");
         $("#selectDate").css("position", "");
         $("#selectDate").css("top", "");
         $("#selectDate").css("left", "");

         $("#txtValue").css("width", "0px");
         $("#txtValue").css("visibility", "hidden");
         $("#txtValue").css("position", "absolute");
         $("#txtValue").css("top", "0");
         $("#txtValue").css("left", "0");

      } else {
         // Hide the logic operator control.
         $("#txtConditionLogic").css("width", "0px");
         $("#txtConditionLogic").css("visibility", "hidden");
         $("#txtConditionLogic").css("position", "absolute");
         $("#txtConditionLogic").css("top", "0px");
         $("#txtConditionLogic").css("left", "0px");

         $("#selectValue").css("width", "0px");
         $("#selectValue").css("visibility", "hidden");
         $("#selectValue").css("position", "absolute");
         $("#selectValue").css("top", "0px");
         $("#selectValue").css("left", "0px");

         $("#selectDate").css("width", "0px");
         $("#selectDate").css("visibility", "hidden");
         $("#selectDate").css("position", "absolute");
         $("#selectDate").css("top", "0");
         $("#selectDate").css("left", "0");

         $("#txtValue").css("width", "100%");
         $("#txtValue").css("visibility", "");
         $("#txtValue").css("position", "");
         $("#txtValue").css("top", "");
         $("#txtValue").css("left", "");
      }

      if ((piDataType == 2) || (piDataType == 4)) {
         // Display the Numeric/Integer operator control.
         $("#selectConditionNum").css("width", "100%");
         $("#selectConditionNum").css("visibility", "");
         $("#selectConditionNum").css("position", "");
         $("#selectConditionNum").css("top", "");
         $("#selectConditionNum").css("left", "");
      }
      else {
         // Hide the Numeric/Integer operator control.
         $("#selectConditionNum").css("width", "0px");
         $("#selectConditionNum").css("visibility", "hidden");
         $("#selectConditionNum").css("position", "absolute");
         $("#selectConditionNum").css("top", "0px");
         $("#selectConditionNum").css("left", "0px");
      }

      if (piDataType == 11) {
         // Display the Date operator control.
         $("#selectConditionDate").css("width", "100%");
         $("#selectConditionDate").css("visibility", "");
         $("#selectConditionDate").css("position", "");
         $("#selectConditionDate").css("top", "");
         $("#selectConditionDate").css("left", "");
      }
      else {
         // Hide the Date operator control.
         $("#selectConditionDate").css("width", "0px");
         $("#selectConditionDate").css("visibility", "hidden");
         $("#selectConditionDate").css("position", "absolute");
         $("#selectConditionDate").css("top", "0px");
         $("#selectConditionDate").css("left", "0px");
      }

      if ((piDataType != -7) && (piDataType != 2) && (piDataType != 4) && (piDataType != 11)) {
         // Display the Character/Working Pattern operator control.
         $("#selectConditionChar").css("width", "100%");
         $("#selectConditionChar").css("visibility", "");
         $("#selectConditionChar").css("position", "");
         $("#selectConditionChar").css("top", "");
         $("#selectConditionChar").css("left", "");
      }
      else {
         // Hide the Character/Working Pattern operator control.
         $("#selectConditionChar").css("width", "0px");
         $("#selectConditionChar").css("visibility", "hidden");
         $("#selectConditionChar").css("position", "absolute");
         $("#selectConditionChar").css("top", "0px");
         $("#selectConditionChar").css("left", "0px");
      }
   }

   function AddToList() {
      var fOK;
      var iDataType;
      var iSize;
      var iDecimals;
      var iTempSize;
      var iTempDecimals;
      var iIndex;
      var sValue;
      var sAddString;
      var sConvertedValue;
      var sDecimalSeparator;
      var sThousandSeparator;
      var sPoint;
      var fDataTypeFound;
      var fSizeFound;
      var fDecimalsFound;

      sDecimalSeparator = '<%:LocaleDecimalSeparator()%>';
      sThousandSeparator = '<%:LocaleThousandSeparator()%>';
      sPoint = ".";

      fOK = false;

      // Determine the data type of the filter column.
      iDataType = 12;
      iSize = 0;
      iDecimals = 0;

      var fieldSelectedVal = $("#selectColumn option:selected").val();
      var fieldSelectedText = $("#selectColumn option:selected").text();
      var dropDown = $("#selectColumn option:selected")[0];
      var iDataType = dropDown.attributes["data-datatype"].value;
      var iSize = dropDown.attributes["data-size"].value;
      var iDecimals = dropDown.attributes["data-decimals"].value;
      if (fieldSelectedVal > 0)
         fOK = true;

      if (fOK == true) {
         sAddString = fieldSelectedText;
         sAddString = sAddString.concat("	");

         if (iDataType == -7) {
            // Logic column (must be the equals operator).
            sAddString = sAddString.concat("equals");
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat($("#selectValue option:selected").text());
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat($("#selectColumn option:selected").val());
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat("1");
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat(iDataType);
         }
      }

      if (fOK == true) {
         if ((iDataType == 2) || (iDataType == 4)) {
            // Numeric/Integer column.
            // Ensure that the value entered is numeric.
            sValue = $("#txtValue").val();
            if (sValue.length == 0) {
               sValue = "0";
            }

            // Convert the value from locale to UK settings for use with the isNaN funtion.
            sConvertedValue = new String(sValue);
            // Remove any thousand separators.
            sConvertedValue = OpenHR.replaceAll(sConvertedValue, sThousandSeparator, "");
            sValue = sConvertedValue;

            // Convert any decimal separators to '.'.
            if ('<%:LocaleDecimalSeparator()%>' != ".") {
               // Existing decimal points are invalid characters.
               sConvertedValue = OpenHR.replaceAll(sConvertedValue, sPoint, "A");
               // Replace the locale decimal marker with the decimal point.
               sConvertedValue = OpenHR.replaceAll(sConvertedValue, sDecimalSeparator, ".");
            }

            if (isNaN(sConvertedValue) == true) {
               fOK = false;
               OpenHR.messageBox("Invalid numeric value entered.");
               $("#txtValue").focus();
            }
            else {
               iIndex = sConvertedValue.indexOf(".");
               if (iDataType == 4) {
                  // Ensure that integer columns are compared with integer values.
                  if (iIndex >= 0) {
                     fOK = false;
                     OpenHR.messageBox("Invalid integer value entered.");
                     $("#txtValue").focus();
                  }
               }
               else {
                  // Ensure numeric columns are compared with numeric values that do not exceed
                  // their defined size and decimals settings.
                  if (iIndex >= 0) {
                     iTempSize = iIndex;
                     iTempDecimals = sConvertedValue.length - iIndex - 1;
                  }
                  else {
                     iTempSize = sConvertedValue.length;
                     iTempDecimals = 0;
                  }

                  if ((sConvertedValue.substr(0, 1) == "+") ||
							(sConvertedValue.substr(0, 1) == "-")) {
                     iTempSize = iTempSize - 1;
                  }

                  if (iTempSize > (iSize - iDecimals)) {
                     fOK = false;
                     OpenHR.messageBox("The column can only be compared to values with " + (iSize - iDecimals) + " digit(s) to the left of the decimal separator.");
                     $("#txtValue").focus();
                  }
                  else {
                     if (iTempDecimals > iDecimals) {
                        fOK = false;
                        OpenHR.messageBox("The column can only be compared to values with " + iDecimals + " decimal place(s).");
                        $("#txtValue").focus();
                     }
                  }
               }

               if (fOK == true) {
                  sAddString = sAddString.concat($("#selectConditionNum option:selected").text());
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat(sValue);
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat($("#selectColumn option:selected").val());
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat($("#selectConditionNum option:selected").val());
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat(iDataType);
               }
            }
         }
      }

      if (fOK == true) {
         if (iDataType == 11) {
            // Date column.
            // Ensure that the value entered is a date.
            sValue = $("#selectDate").val();

            if (sValue.length == 0) {
               sAddString = sAddString.concat($("#selectConditionDate option:selected").text());
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat("");
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat($("#selectColumn option:selected").val());
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat($("#selectConditionDate option:selected").val());
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat(iDataType);
            }
            else {
               // Convert the date to SQL format (use this as a validation check).
               // An empty string is returned if the date is invalid.
               if (OpenHR.convertLocaleDateToSQL(sValue) == "null") {
                  fOK = false;
                  OpenHR.messageBox("Invalid date value entered.");
                  $("#txtValue").focus();
               }
               else {
                  //						sValue = OpenHR.convertLocaleDateToSQL(sValue);

                  sAddString = sAddString.concat($("#selectConditionDate option:selected").text());
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat(sValue);
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat($("#selectColumn option:selected").val());
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat($("#selectConditionDate option:selected").val());
                  sAddString = sAddString.concat("	");
                  sAddString = sAddString.concat(iDataType);
               }
            }
         }
      }

      if (fOK == true) {
         if ((iDataType != -7) && (iDataType != 2) && (iDataType != 4) && (iDataType != 11)) {
            // Character/Working Pattern column.
            sValue = $("#txtValue").val();

            sAddString = sAddString.concat($("#selectConditionChar option:selected").text());
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat(sValue);
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat($("#selectColumn option:selected").val());
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat($("#selectConditionChar option:selected").val());
            sAddString = sAddString.concat("	");
            sAddString = sAddString.concat(iDataType);
         }
      }

      if (fOK == true) {
         var items = sAddString.split("\t");
         $("#DBGridFilterRecords").addRowData(
					$("#DBGridFilterRecords").getGridParam("reccount") + 1, //ID
					{ //Data
					   'FieldName': items[0],
					   'OperatorName': items[1],
					   'FilterValue': $.jgrid.htmlEncode(items[2]),
					   'FieldID': items[3],
					   'OperatorID': items[4],
					   'FieldDataType': items[5],
                  'ID':0
					},
					'last'); //Add the record at the end

         //Select the newly added record
         $("#DBGridFilterRecords").jqGrid('setSelection', $("#DBGridFilterRecords").getGridParam("reccount"));

         FilterSelect_refreshControls();
         $("select#selectColumn").focus();
         $("select#selectColumn")[0].selectedIndex = 0;
         $("#txtValue").val('');
         var resetdropDown = $("select#selectColumn option:selected")[0];
         var iResetDataType = resetdropDown.attributes["data-datatype"].value;
         refreshSelectColumnCombo();         
         if (iResetDataType == -7)
         {
            $("select#selectValue")[0].selectedIndex = 0;
         }
         else if ((iResetDataType == 2) || (iDataType == 4))
         {
            $("select#selectConditionNum")[0].selectedIndex = 0;
         }
         else if (iResetDataType == 11) // Date column.
         {
            $("select#selectConditionDate")[0].selectedIndex = 0;
         } 
         else if ((iResetDataType != -7) && (iResetDataType != 2) && (iResetDataType != 4) && (iResetDataType != 11))  // Character/Working Pattern column.
         {
            $("select#selectConditionChar")[0].selectedIndex = 0;
         }
         
         enableSaveButton();
      }
   }

   function FilterSelect_removeAll() {
      $("#DBGridFilterRecords").jqGrid('clearGridData');
      FilterSelect_refreshControls();
      enableSaveButton();
   }

   function FilterSelect_remove() {

      if ($("#DBGridFilterRecords").getGridParam("reccount") > 0) {
         var iRowIndex = $("#DBGridFilterRecords").jqGrid('getCell', $('#DBGridFilterRecords').jqGrid('getGridParam', 'selrow'), 5); //5 -> ID

         if (($("#DBGridFilterRecords").getGridParam("reccount") == 1) && (iRowIndex == 0)) {
            FilterSelect_removeAll();
         }
         else {
            var grid = $("#DBGridFilterRecords");
            var myDelOptions = {
               // because I use "local" data I don't want to send the changes
               // to the server so I use "processing:true" setting and delete
               // the row manually in onclickSubmit
               onclickSubmit: function (options) {
                  var grid_id = $.jgrid.jqID(grid[0].id),
								grid_p = grid[0].p,
								newPage = grid_p.page,
								rowids = grid_p.multiselect ? grid_p.selarrrow : [grid_p.selrow];

                  // reset the value of processing option which could be modified
                  options.processing = true;

                  // delete the row
                  $.each(rowids, function () {
                     grid.delRowData(this);
                  });
                  $.jgrid.hideModal("#delmod" + grid_id,
															{
															   gb: "#gbox_" + grid_id,
															   jqm: options.jqModal, onClose: options.onClose
															});
                  return true;
               },
               processing: true
            };

            grid.jqGrid('delGridRow', grid.jqGrid('getGridParam', 'selarrrow'), myDelOptions);

            $("#dData").click(); //To remove the "delete confirmation" dialog
         }
      }

      FilterSelect_refreshControls();
      enableSaveButton();
   }



   function FilterSelect_refreshControls() {     
     
      if ($("#DBGridFilterRecords").getGridParam("reccount") > 0) {
         button_disable($("#btnRemoveAll")[0], false);

         if ($('#DBGridFilterRecords').jqGrid('getGridParam', 'selrow') != null && $('#DBGridFilterRecords').jqGrid('getGridParam', 'selrow').length > 0) {                      
            button_disable($("#btnRemove")[0], false);
         }
         else {
            button_disable($("#btnRemove")[0], true);
         }
      }
      else {
         button_disable($("#btnRemoveAll")[0], true);
         button_disable($("#btnRemove")[0], true);
      }

      var topID = $("#DBGridFilterRecords").getDataIDs()[0]
      $("#DBGridFilterRecords").jqGrid("setSelection", topID);
   }

   function GetExistingFilterDefinition(sFilterDef)
   {
      var iColumnID;
      var iOperatorID;
      var sValue;
      var sColumnName;
      var iColumnDataType;
      var fFilterOK = false;
      var sOperatorName;
      var fFound;

      var ar = $.parseJSON(sFilterDef);
      if (ar.rows && ar.rows.length > 0) {

         for (i = 0; i < ar.rows.length; i++) {

            iColumnID = ar.rows[i].FieldID;
            sColumnName = ar.rows[i].FieldName;
            iOperatorID = ar.rows[i].OperatorID;
            sValue = ar.rows[i].FilterValue                        
            sValue = sValue.replace(/\*ALL/g, '*');
            iColumnDataType = ar.rows[i].FieldDataType;

            fFilterOK = true;

            if (fFilterOK == true) {
               // Get the operator name.
               fFound = false;
               if (iOperatorID == 1) {
                  fFound = true;
                  sOperatorName = "is equal to";
               }
               if (iOperatorID == 2) {
                  fFound = true;
                  sOperatorName = "is NOT equal to";
               }
               if (iOperatorID == 3) {
                  fFound = true;
                  if (iColumnDataType == 11) {
                     sOperatorName = "is equal to or before";
                  } else {
                     sOperatorName = "is less than or equal to";
                  }
               }
               if (iOperatorID == 4) {
                  fFound = true;
                  if (iColumnDataType == 11) {
                     sOperatorName = "is equal to or after";
                  } else {
                     sOperatorName = "is greater than or equal to";
                  }
               }
               if (iOperatorID == 5) {
                  fFound = true;
                  if (iColumnDataType == 11) {
                     sOperatorName = "after";
                  } else {
                     sOperatorName = "is greater than";
                  }
               }
               if (iOperatorID == 6) {
                  fFound = true;
                  if (iColumnDataType == 11) {
                     sOperatorName = "before";
                  } else {
                     sOperatorName = "is less than";
                  }
               }
               if (iOperatorID == 7) {
                  fFound = true;
                  sOperatorName = "contains";
               }
               if (iOperatorID == 8) {
                  fFound = true;
                  sOperatorName = "does not contain";
               }
               if (fFound == false) {
                  fFilterOK = false;
               }
            }

            if (fFilterOK == true) {
               // Add the filter definition to the grid.
               sAddString = sColumnName;
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat(sOperatorName);
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat(sValue);
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat(iColumnID);
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat(iOperatorID);
               sAddString = sAddString.concat("	");
               sAddString = sAddString.concat(iColumnDataType);

               //Determine if the grid already exists...
               if ($("#DBGridFilterRecords").getGridParam("reccount") == undefined) { //It doesn't exist, create it
                  $("#DBGridFilterRecords").jqGrid({
                     autoencode: true,
                     multiselect: false,
                     data: colData,
                     datatype: 'local',
                     colNames: colNames,
                     colModel: colMode,
                     rowNum: 1000,
                     autowidth: true,
                     shrinkToFit: true,
                     onSelectRow: function () {
                        button_disable(frmFilterForm.cmdRemoveAll, false);
                        button_disable(frmFilterForm.cmdRemove, false);
                     },
                     editurl: 'clientArray'
                  }).jqGrid('hideCol', 'cb');
               }

               var items = sAddString.split("\t");
               $("#DBGridFilterRecords").addRowData(
                     $("#DBGridFilterRecords").getGridParam("reccount") + 1, //ID
                     { //Data
                        'FieldName': items[0],
                        'OperatorName': items[1],
                        'FilterValue': $.jgrid.htmlEncode(items[2]),
                        'FieldID': items[3],
                        'OperatorID': items[4],
                        'FieldDataType': items[5],
                        'ID': 0
                     },
                     'last'); //Add the record at the end

               //Select the newly added record
               $("#DBGridFilterRecords").jqGrid('setSelection', $("#DBGridFilterRecords").getGridParam("reccount"));

               FilterSelect_refreshControls();
            }
         }

      }

      //Select the top filter record
      if ($("#DBGridFilterRecords").getGridParam("reccount") > 0) {
         $("#DBGridFilterRecords").jqGrid('setSelection', 1);
      }

   }
</script>


