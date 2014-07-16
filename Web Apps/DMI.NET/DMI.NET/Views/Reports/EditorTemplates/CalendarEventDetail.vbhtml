@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.CalendarEventDetailViewModel)

@Html.BeginForm("PostCalendarEvent", "Reports", FormMethod.Post, New With {.id = "frmPostCalendarEvent"}))

<div class="left">

	@Html.HiddenFor(Function(model) model.EventKey)
	@Html.HiddenFor(Function(model) model.CalendarReportID)
	@Html.HiddenFor(Function(model) model.FilterHidden)

	<fieldset>
		<legend>Event :</legend>

		@Html.DisplayNameFor(Function(model) model.Name)
		@Html.EditorFor(Function(model) model.Name)
		<br />
		@Html.LabelFor(Function(model) model.TableID)
		@Html.TableDropdown("CalendarEventTableID", Model.TableID, Model.AvailableTables, "changeEventTable(event);")
		<br/>
		<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
		@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtEventFilter", .readonly = "true"})
		@Html.EllipseButton("cmdBasePicklist", "selectRecordOption('event', 'filter')", True)

	</fieldset>


	<fieldset>
		<legend>Event Start :</legend>

		@Html.ColumnDropdown2("EventStartDateID", Model.EventStartDateID)

		@Html.EditorFor(Function(model) model.EventStartSessionID)
	</fieldset>

	<fieldset>
		<legend>Event End :</legend>

		@Html.EditorFor(Function(model) model.EventEndType)
		@Html.EditorFor(Function(model) model.EventEndDateID)
		@Html.EditorFor(Function(model) model.EventEndSessionID)
		@Html.EditorFor(Function(model) model.EventDurationID)

	</fieldset>

</div>

<div class="right">

	<fieldset>
		<legend>Key</legend>

		@Html.DisplayFor(Function(model) model.LegendType)
		@Html.DisplayFor(Function(model) model.LegendCharacter)
		@Html.DisplayFor(Function(model) model.LegendLookupTableID)
		@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)
		@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)
		@Html.DisplayFor(Function(model) model.LegendLookupCodeID)
		@Html.DisplayFor(Function(model) model.LegendEventColumnID)

</fieldset>

	<fieldset>
		<legend>Event Description</legend>

		@Html.ColumnDropdown2("EventDesc1ColumnID", Model.EventDesc1ColumnID)
		@Html.ColumnDropdown2("EventDesc2ColumnID", Model.EventDesc2ColumnID)

</fieldset>



	<dl>

		<dt>
			@Html.DisplayNameFor(Function(model) model.FilterHidden)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.FilterHidden)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.FilterName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.FilterName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.EventStartSessionName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.EventStartSessionName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.EventEndDateName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.EventEndDateName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.EventEndSessionName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.EventEndSessionName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.EventDurationName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.EventDurationName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.LegendTypeName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.LegendTypeName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.EventDesc1ColumnName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.EventDesc1ColumnName)
		</dd>

		<dt>
			@Html.DisplayNameFor(Function(model) model.EventDesc2ColumnName)
		</dt>

		<dd>
			@Html.DisplayFor(Function(model) model.EventDesc2ColumnName)
		</dd>

	</dl>
</div>

<input type="button" value="OK" onclick="postThisCalendarEvent();" />

@code
	Html.EndForm()
End Code



<script type="text/javascript">

	function changeEventTable(event) {

		// Reload dropdowns from server

		debugger;



	}

	function postThisCalendarEvent() {

		//var datarow = {
		//	TableID: $("#TableID").val(),
		//	FilterID: $("#txtChildFilterID").val(),
		//	OrderID: $("#txtChildFieldOrderID").val(),
		//	TableName: $("#txtChildTableID").val(),
		//	FilterName: $("#txtChildFilter").val(),
		//	OrderName: $("#txtChildOrder").val(),
		//	Records: $("#txtChildRecords").val()
		//};

		//// Update client
		//var su = jQuery("#ChildTables").jqGrid('addRowData', 99, datarow);

		// Post to server
		var frmSubmit = $("#frmPostCalendarEvent");
		OpenHR.postForm(frmSubmit);


		$("#divPopupGetChildTable").dialog("close");
		$("#divPopupGetChildTable").empty();

	}


</script>