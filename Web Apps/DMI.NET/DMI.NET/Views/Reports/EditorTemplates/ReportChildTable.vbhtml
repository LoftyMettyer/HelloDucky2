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
		@Html.TextBoxFor(Function(m) m.OrderName, New With {.id = "txtFieldRecOrder", .readonly = "true"})
		@Html.EllipseButton("cmdBaseFilter", "selectRecordOrder()", True)
		@<br />
		@Html.LabelFor(Function(m) m.Records)
		@Html.TextBoxFor(Function(m) m.Records, New With {.id = "txtChildRecords"})
		@<br />
		@<input type="button" value="OK" onclick="postThisChildTable();" />
		@<input type="button" value="Cancel" onclick="closeThisChildTable();" />

		
	End Using 
</fieldset>

<script>

	function selectRecordOrder() {
		var sURL;

		sURL = "fieldRec" +
				"?selectionType=" + "ORDER" +
				"&txtTableID=" + $("#ChildTableID option:selected").val() +
				"&selectedID=" + $("#txtChildFieldOrderID").val();
		openDialog(sURL, (screen.width) / 3 + 40, (screen.height) / 2 - 30, "no", "no");
	}

	function closeThisChildTable() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

		function postThisChildTable() {

			var datarow = {
				ReportID: '@Model.ReportID',
				ReportType: '@Model.ReportType',
				TableID: $("#ChildTableID").val(),
				FilterID: $("#txtChildFilterID").val(),
				OrderID: $("#txtChildFieldOrderID").val(),
				TableName: $("#ChildTableID option:selected").text(),
				FilterName: $("#txtChildFilter").val(),
				OrderName: $("#txtFieldRecOrder").val(),
				Records: $("#txtChildRecords").val()
			};

			// Update client
			$('#ChildTables').jqGrid('delRowData', $("#ChildTableID").val())
			var su = jQuery("#ChildTables").jqGrid('addRowData', $("#ChildTableID").val(), datarow);

			// Post to server
			OpenHR.postData("Reports/PostChildTable", datarow, loadAvailableTablesForReport)

			$("#divPopupReportDefinition").dialog("close");
			$("#divPopupReportDefinition").empty();

		}

</script>