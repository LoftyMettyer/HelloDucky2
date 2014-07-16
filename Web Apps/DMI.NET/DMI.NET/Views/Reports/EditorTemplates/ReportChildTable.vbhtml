@Imports DMI.NET.Classes
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ReportChildTables)

<fieldset>
	@Using (Html.BeginForm("PostChildTable", "Reports", FormMethod.Post, New With {.id = "frmPostChildTable"}))

 		@Html.HiddenFor(Function(m) m.ID)

	 	@Html.LabelFor(Function(m) m.TableID) 
	 	@Html.TableDropdown("TableID", Model.TableID, Model.AvailableTables, Nothing)
	 
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
				TableID: $("#TableID").val(),
				FilterID: $("#txtChildFilterID").val(),
				OrderID: $("#txtChildFieldOrderID").val(),
				TableName: $("#txtChildTableID").val(),
				FilterName: $("#txtChildFilter").val(),
				OrderName: $("#txtChildOrder").val(),
				Records: $("#txtChildRecords").val()
			};

			// Update client
			var su = jQuery("#ChildTables").jqGrid('addRowData', 99, datarow);

			// Post to server
			var frmSubmit = $("#frmPostChildTable");
			OpenHR.postForm(frmSubmit);

			$("#divPopupGetChildTable").dialog("close");
			$("#divPopupGetChildTable").empty();

		}

</script>