@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ChildTableViewModel)

<fieldset>
	@Using (Html.BeginForm("PostChildTable", "Reports", FormMethod.Post, New With {.id = "frmPostChildTable"}))

 		@Html.HiddenFor(Function(m) m.ReportID)

	 	@Html.LabelFor(Function(m) m.TableID) 
	 	@Html.TableDropdown("TableID", "ChildTableID", Model.TableID, Model.AvailableTables, Nothing)
	 
		@<br/>
		@Html.HiddenFor(Function(m) m.FilterID, New With {.id = "txtChildFilterID"})
		@Html.LabelFor(Function(m) m.FilterName)
		@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtChildFilter", .readonly = "true"})
		@Html.EllipseButton("cmdBaseFilter", "selectRecordOption('child', 'filter')", True)
		@<br />
		@Html.LabelFor(Function(m) m.OrderName)
		@Html.HiddenFor(Function(m) m.OrderID, New With {.id = "txtChildFieldOrderID"})
		@Html.TextBoxFor(Function(m) m.OrderName, New With {.id = "txtChildOrder", .readonly = "true"})
		@Html.EllipseButton("cmdBaseFilter", "selectRecordOrder()", True)
		@<br />
		@Html.LabelFor(Function(m) m.Records)
		@Html.TextBoxFor(Function(m) m.Records, New With {.id = "txtChildRecords"})
		@<br />
		@<input type="button" value="OK" onclick="postThisChildTable();" />
		
	End Using 
</fieldset>

<script>

		function postThisChildTable() {

			var datarow = {
				ReportID: '@Model.ReportID',
				TableID: $("#ChildTableID").val(),
				FilterID: $("#txtChildFilterID").val(),
				OrderID: $("#txtChildFieldOrderID").val(),
				TableName: $("#ChildTableID option:selected").text(),
				FilterName: $("#txtChildFilter").val(),
				OrderName: $("#txtChildOrder").val(),
				Records: $("#txtChildRecords").val()
			};

			// Update client
			$('#ChildTables').jqGrid('delRowData', $("#ChildTableID").val())
			var su = jQuery("#ChildTables").jqGrid('addRowData', $("#ChildTableID").val(), datarow);

			// Update available tables
			loadRelatedTables();

			// Post to server
			OpenHR.postData("Reports/PostChildTable", datarow)

			$("#divPopupReportDefinition").dialog("close");
			$("#divPopupReportDefinition").empty();

		}

</script>