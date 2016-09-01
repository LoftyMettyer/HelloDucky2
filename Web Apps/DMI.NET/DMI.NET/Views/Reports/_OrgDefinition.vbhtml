@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of OrganisationReportModel)
@Html.HiddenFor(Function(m) m.ID, New With {.id = "txtReportID"})
@Html.HiddenFor(Function(m) m.Timestamp)
@Html.HiddenFor(Function(m) m.ReportType, New With {.id = "txtReportType"})
@Html.HiddenFor(Function(m) m.IsReadOnly)
@Html.HiddenFor(Function(m) m.ActionType)
@Html.HiddenFor(Function(m) m.ValidityStatus)

   <div class="width100">
      <fieldset class="floatleft width50 bordered">
         <legend class="fontsmalltitle">Identification :</legend>

         <fieldset class="">
            @Html.LabelFor(Function(m) m.Name)
            <div class="width70 floatright">
               @Html.TextBoxFor(Function(m) m.Name, New With {.class = "width100 floatright", .maxlength = 50})
               @Html.ValidationMessageFor(Function(m) m.Name)
            </div>
         </fieldset>

         <fieldset class="">
            @Html.LabelFor(Function(m) m.CategoryList)
            @Html.DropDownListFor(Function(m) m.CategoryID, New SelectList(Model.CategoryList, "Value", "Text"), New With {.class = "width70 floatright", .onchange = "enableSaveButton()"})
         </fieldset>

         <fieldset class="">
            @Html.LabelFor(Function(m) m.Description)
            <div id="textareadescription" class="width70 floatright">
               @Html.TextArea("description", Model.Description, New With {.class = "width100 floatright"})
               @Html.ValidationMessageFor(Function(m) m.Description)
            </div>
         </fieldset>
      </fieldset>

      <fieldset id="DataRecordsPermissions" class="floatleft overflowhidden width50">
         <legend class="fontsmalltitle">Data :</legend>
         <div class="inner">
            <fieldset id="BaseViewText" class="">
               <span>Base View :</span>
               @Html.DropDownListFor(Function(m) m.BaseViewID, New SelectList(Model.BaseViewList, "Id", "Name"), New With {.class = "width70 floatright", .id = "BaseViewId", .name = "BaseViewId", .onchange = "requestChangeReportBaseView()"})
               <input type="hidden" id="OriginalBaseViewId" />
               <input type="hidden" id="IsBaseViewChange" value="False" />
            </fieldset>

            <input type="hidden" id="ctl_DefinitionChanged" name="HasChanged" value="false" />
            <input type="hidden" id="baseHidden" name="baseHidden">
         </div>
      </fieldset>

      <fieldset id="AccessPermissions" class="table">
         <legend class="fontsmalltitle">Group Access :</legend>

         <fieldset>
            <div class="nowrap tablelayout" style="margin-top: 0;">
               <div class="tablerow">
                  <label>Owner :</label>
                  @Html.TextBoxFor(Function(m) m.Owner, New With {.readonly = "true"})
               </div>
               <br />
               <div class="tablerow">
                  <label>Access :</label>
                  @Html.AccessGrid("GroupAccess", Model.GroupAccess, Model.IsGroupAccessHiddenWhenCopyTheDefinition, New With {.id = "tblGroupAccess"})
                  <input type="hidden" id="IsForcedHidden" />
               </div>
            </div>
         </fieldset>
      </fieldset>
   </div>

   <script type="text/javascript">

      $(document).ajaxStop(function () {
         $('#description').removeAttr('style');
         $('#Name').removeAttr('style');
      });

      $(function () {

         HideToolsButtons();

         $('fieldset').css("border", "0");
         $('table').css("border", "0");

         refreshViewAccess();

         tableToGrid('#tblGroupAccess', {
            autoWidth: true, height: 150, cmTemplate: { sortable: false },
            afterInsertRow: function (rowid, aData) {
               // set empty tooltip for access dropdown
               $("#tblGroupAccess").setCell(rowid, 'Access', '', '', { title: '' })
            }
         });

         menu_toolbarEnableItem('mnutoolSaveReport', false);

         if ($("#ActionType").val() == '@UtilityActionType.Copy') {
            enableSaveButton()
         }

         if (isDefinitionReadOnly()) {
            $("#frmReportDefintion input").prop('disabled', "disabled");
            $("#frmReportDefintion textarea").prop('disabled', "disabled");
            $("#frmReportDefintion select").prop('disabled', "disabled");
            $("#frmReportDefintion :button").prop('disabled', "disabled");
         }
         else {
            $("#frmReportDefintion input").on("keydown", function () { enableSaveButton(); });
            $("#frmReportDefintion textarea").on("keydown", function () { enableSaveButton(); });
            $("#frmReportDefintion input").on("change", function () { enableSaveButton(); });

            //bind click event on the css class for the button and change event for the dropdown to enable the save button
            $("#frmReportDefintion .enableSaveButtonOnClick").on("click", function () { enableSaveButton(); });
            $("#frmReportDefintion .enableSaveButtonOnComboChange").on("change", function () { enableSaveButton(); });
         }
      });

      function isDefinitionReadOnly() {
         return ($("#IsReadOnly").val() == "True");
      }

      function setAllSecurityGroups() {

         var setTo = $("#drpSetAllSecurityGroups").val();
         if (setTo.length > 0) $(".reportViewAccessGroup").val(setTo);

      }

      function refreshViewAccess() {

         var bViewAccessEnabled = true;
         var list;

         $(".reportViewAccessGroup").prop('disabled', false);
         $("#drpSetAllSecurityGroups").prop('disabled', false);
         $(".reportViewAccessGroup").removeClass('ui-state-disabled');

         $(".ViewAccess").each(function (index) {
            if ((this).innerText == "HD" || (this).value == "HD") {
               bViewAccessEnabled = false;
            }
         });

         if (!bViewAccessEnabled) {
            $("#IsForcedHidden").val(true);
            $(".reportViewAccessGroup").prop('disabled', true);
            $("#drpSetAllSecurityGroups").prop('disabled', true);
            $(".reportViewAccessGroup").addClass('ui-state-disabled');
         }
      }

      function requestChangeReportBaseView(target) {

         // Get count : Check if any column is selected
         // Get count : Check if any filer is selected
         var filterCount = $("#DBGridFilterRecords").getGridParam("reccount");
         var columnCount = $("#SelectedColumns").getGridParam("reccount");

         if (filterCount > 0 || columnCount > 0) {
            OpenHR.modalPrompt("Changing the Base View will reset definition data.<br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
               if (answer == 6) { // Yes
                  changeBaseView();
               }
               else {
                  $('#BaseViewId')[0].selectedIndex = $("#OriginalBaseViewId").val();
               }
            });
         }
         else {
            changeBaseView();
         }
      }

      function changeBaseView() {
         // Post base table change to server
         var dataSend = {
            ReportID: '@Model.ID',
            ReportType: '@Model.ReportType',
            BaseViewId: $("#BaseViewId option:selected").val(),
            __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
         };

         OpenHR.postData("Reports/ChangeBaseView", dataSend, changeReportBaseViewCompleted);
         $("#OriginalBaseViewId").val($('#BaseViewId')[0].selectedIndex);
         $("#SelectedTableID").val($("#BaseViewId option:selected").val());
         GetFilterColumns();
      }

      function changeReportBaseViewCompleted(json) {

         // Remove filters.
         FilterSelect_removeAll();

         // Remove Columns.
         removeAllSelectedColumns(true);

         // Enables save button
         enableSaveButton();
      }

      //This function will be used to select the first row of provided grid control
      function selectGridTopRow(gridControl) {
         // Highlight top row of selected columns grid
         ids = gridControl.jqGrid("getDataIDs");
         if (ids && ids.length > 0)
            gridControl.jqGrid("setSelection", ids[0]);
      }

      function enableSaveButton() {

         if (!isDefinitionReadOnly()) {
            $("#ctl_DefinitionChanged").val("true");
            menu_toolbarEnableItem('mnutoolSaveReport', true);
         }
      }

      function saveReportDefinition(prompt) {

         var bHasChanged = $("#ctl_DefinitionChanged").val();

         if (prompt == true) {
            if (bHasChanged == "true") {

               OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
                  if (answer == 1) {
                     validateReportDefinition();
                  }
               });
            }
            else {
               return 6;
            }

         } else {
            validateReportDefinition()
         }

         return 0;
      }

     
      function validateReportDefinition() {

         var gridData;
         var gridFitlerData;
         menu_ShowWait("Saving...");
         //Filters selected
         gridFitlerData = $("#DBGridFilterRecords").getRowData();
         $('#txtFilterColumns').val(JSON.stringify(gridFitlerData));
        
         // Columns selected
         gridData = $("#SelectedColumns").getRowData();
         $('#txtCSAAS').val(JSON.stringify(gridData));

         var $form = $("#frmReportDefintion");
         $(".reportViewAccessGroup").prop('disabled', false);
         $("#drpSetAllSecurityGroups").prop('disabled', false);
         $(".reportViewAccessGroup").removeClass('ui-state-disabled');

         $.ajax({
            url: $form.attr("action"),
            type: $form.attr("method"),
            data: $form.serialize(),
            async: true,
            error: function (json) {
               OpenHR.modalPrompt("Invalid characters in report definition.", 0, "OpenHR");
            },
            success: function (json) {

               switch (json.ErrorCode) {
                  case 0:
                     submitReportDefinition();
                     break;

                  case 1:
                     OpenHR.modalPrompt(json.ErrorMessage, 0, "OpenHR");
                     break;

                  case -1:
                     OpenHR.modalPrompt(json.ErrorMessage, 0, "OpenHR");
                     break;

                  default:
                     OpenHR.modalPrompt(json.ErrorMessage, 4, "Confirm").then(function (answer) {
                        if (answer == 6) {
                           submitReportDefinition();
                        }
                     });
                     break;

               }
               refreshViewAccess();
            }
         });
      }

      function submitReportDefinition() {
         $("#ValidityStatus").val('ServerCheckComplete');
         $(".reportViewAccessGroup").prop('disabled', false);
         $("#drpSetAllSecurityGroups").prop('disabled', false);
         $(".reportViewAccessGroup").removeClass('ui-state-disabled');
         var frmSubmit = $("#frmReportDefintion")[0];
         OpenHR.submitForm(frmSubmit);
      }

      function cancelReportDefinition() {

         var bHasChanged = $("#ctl_DefinitionChanged").val();

         if (bHasChanged == "true") {
            OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
               if (answer == 1) {  // OK
                  menu_loadDefSelPage('@CInt(Model.ReportType)', '@Model.ID', $("#BaseViewId option:selected").val(), true);
                  return 6;
               }
            })
         }
         else {

            menu_loadDefSelPage('@CInt(Model.ReportType)', '@Session("utilid")', $("#BaseViewId option:selected").val(), true);
         }

         return false;
      }

      //If the Base Table is Personnel Records then 'Include Bank Holidays', 'Working Days Only' and 'Show Bank Holidays' should enable.
      function disableEnableWorkingDaysOrHolidays(bDisabled) {
         $('#IncludeBankHolidays').prop('disabled', bDisabled);
         $('#WorkingDaysOnly').prop('disabled', bDisabled);
         $('#ShowBankHolidays').prop('disabled', bDisabled);

         if (bDisabled) {
            $("#label_IncludeBankHolidays").css('opacity', '0.5');
            $("#label_WorkingDaysOnly").css('opacity', '0.5');
            $("#label_ShowBankHolidays").css('opacity', '0.5');
         }
         else {
            $("#label_IncludeBankHolidays").css('opacity', '1');
            $("#label_WorkingDaysOnly").css('opacity', '1');
            $("#label_ShowBankHolidays").css('opacity', '1');
         }
      }

      // Show/Hide tools buttons
      function HideToolsButtons() {

         // Set the picklist & filter ribbon button to visible.
         menu_setVisibleMenuItem("mnutoolPicklistReport", false);
         menu_setVisibleMenuItem("mnutoolFilterReport", false);
         menu_setVisibleMenuItem("mnutoolCalculationReport", false);
      }

      // Enable/Disable the save of report definition button.
      // (E.g. When the user comes to tools screen from the report definition and modify the tools definition. Then, even the report defition is not modified, the Save button for the report definition would remain enabled in this case.)
      function EnableDisableSaveButton() {
         if (!isDefinitionReadOnly()) {
            var bHasChanged = $("#ctl_DefinitionChanged").val();
            menu_toolbarEnableItem('mnutoolSaveReport', (bHasChanged == "true") ? true : false);
         }
         else {
            menu_toolbarEnableItem('mnutoolSaveReport', false);
         }
      }


   </script>


