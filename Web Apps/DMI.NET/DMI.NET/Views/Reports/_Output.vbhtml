@Imports DMI.NET
@Imports DMI.NET.Helpers

@code
	Html.EnableClientValidation()
End Code

@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportOutputModel)
<fieldset>

	<fieldset class="border0 width20 floatleft">
		<legend class="fontsmalltitle">Output Formats</legend>
		<fieldset id="outputformats">
			@Html.RadioButton("Output.Format", 0, Model.Format = OutputFormats.DataOnly, New With {.onchange = "changeOutputType('DataOnly')"})
			Data Only
			<br />

			<div class="hideforcalendarreport">
				@Html.RadioButton("Output.Format", 1, Model.Format = OutputFormats.CSV, New With {.onchange = "changeOutputType('CSV')"})
				<span class="DataManagerOnly">CSV File</span>
				<br />
			</div>

			@Html.RadioButton("Output.Format", 2, Model.Format = OutputFormats.HTML, New With {.onchange = "changeOutputType('HTML')"})
			<span class="DataManagerOnly">HTML Document</span>
			<br />

			@Html.RadioButton("Output.Format", 3, Model.Format = OutputFormats.WordDoc, New With {.onchange = "changeOutputType('WordDoc')"})
			<span class="DataManagerOnly">Word Document</span>
			<br />

			@Html.RadioButton("Output.Format", 4, Model.Format = OutputFormats.ExcelWorksheet, New With {.onchange = "changeOutputType('ExcelWorksheet')"})
			Excel Worksheet
			<br />

			<div class="hideforcalendarreport">
				@Html.RadioButton("Output.Format", 5, Model.Format = OutputFormats.ExcelGraph, New With {.onchange = "changeOutputType('ExcelGraph')"})
				Excel Chart
				<br />
				@Html.RadioButton("Output.Format", 6, Model.Format = OutputFormats.ExcelPivotTable, New With {.onchange = "changeOutputType('ExcelPivotTable')"})
				Excel Pivot Table
			</div>
			<br />
		</fieldset>
	</fieldset>

	<fieldset id="outputdestinatonfieldset" class="border0 floatleft width70">
		<legend class="fontsmalltitle">Output Destinations</legend>

		<fieldset class="border0 reportdefpreview">
			<div>
				@Html.CheckBoxFor(Function(m) m.IsPreview, New With {Key .Name = "Output.IsPreview"})
				@Html.LabelFor(Function(m) m.IsPreview)
			</div>
		</fieldset>

		<fieldset class="border0 reportdefscreen">
			<div>
				@Html.CheckBoxFor(Function(m) m.ToScreen, New With {Key .Name = "Output.ToScreen"})
				@Html.LabelFor(Function(m) m.ToScreen)
			</div>
		</fieldset>

		<fieldset class="border0 reportdefprinter DataManagerOnly">
			<div class="width30 floatleft">
				@Html.CheckBoxFor(Function(m) m.ToPrinter, New With {Key .Name = "Output.ToPrinter"})
				@Html.LabelFor(Function(m) m.ToPrinter, New With {.class = "DataManagerOnly"})
			</div>
			<div class="width70 floatleft">
				@Html.TextBoxFor(Function(m) m.PrinterName, New With {.Name = "Output.PrinterName", .placeholder = "Default Printer", .class = "DataManagerOnly width100", .readonly = "true"})
			</div>
		</fieldset>

		<fieldset class="border0 reportdeffile">
			<div class="">
				<div class="width30 floatleft">
					@Html.CheckBoxFor(Function(m) m.SaveToFile, New With {Key .Name = "Output.SaveToFile", .onclick = "setOutputToFile();", .class = "DataManagerOnly"})
					@Html.LabelFor(Function(m) m.SaveToFile, New With {.class = "DataManagerOnly"})
				</div>
				<div class="width70 floatleft">
					@Html.TextBoxFor(Function(m) m.Filename, New With {.Name = "Output.Filename", .placeholder = "File Name", .readonly = "true", .class = "width100"})
				</div>
			</div>
		</fieldset>

		<fieldset>
			<div id="outputtabfilenametextbox" class="">
				@Html.LabelFor(Function(m) m.SaveExisting)
				@Html.EnumDropDownListFor(Function(m) m.SaveExisting, New With {.class = "DataManagerOnly", .readonly = "true"})
			</div>
		</fieldset>

		<fieldset id="fieldsetsendemail" class="reportdefemail">
			<div class="">
				<div class="width30 floatleft">
					@Html.CheckBoxFor(Function(m) m.SendToEmail, New With {Key .Name = "Output.SendToEmail", .onclick = "setOutputToEmail();"})
					@Html.LabelFor(Function(m) m.SendToEmail)
				</div>
				<div class="width70 floatleft">
					@Html.HiddenFor(Function(m) m.EmailGroupID, New With {.Name = "Output.EmailGroupID", .id = "txtEmailGroupID"})
					@Html.TextBoxFor(Function(m) m.EmailGroupName, New With {.Name = "Output.EmailGroupName", .id = "txtEmailGroup", .readonly = "readonly", .class = "display-textbox-emails", .style = ""})
					<input type="button" class="" id="cmdEmailGroup" name="cmdEmailGroup" value="..." onclick="selectEmailGroup()" />
				</div>
			</div>
		</fieldset>

		<fieldset id="fieldsetsubjectemail">			
				@Html.LabelFor(Function(m) m.EmailSubject, New With {.class = "display-label_emails"})
				@Html.TextBox("Output.EmailSubject", Model.EmailSubject, New With {.class = "display-textbox-emails"})			
			</fieldset>
		<fieldset id="fieldseattachas">
			@Html.LabelFor(Function(m) m.EmailAttachmentName, New With {.class = "display-label_emails"})
			@Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {.Name = "Output.EmailAttachmentName", .class = "display-textbox-emails"})
		</fieldset>
		<br />
		@Html.ValidationMessage("Output.EmailGroupID")		<br />
		@Html.ValidationMessage("Output.EmailSubject")		<br />
		@Html.ValidationMessage("Output.EmailAttachmentName")		<br />
		@Html.ValidationMessage("Output.FileName")		<br />
	</fieldset>

	<fieldset class="DataManagerOnly width100">
		Note: Options marked in red are unavailable in OpenHR Web.
	</fieldset>
</fieldset>

<script type="text/javascript">

	function setOutputToFile() {

		var bSelected = $("#SaveToFile").prop('checked');
		$(".reportdeffile").children().attr("readonly", !bSelected);

		if (!bSelected) {
			$(".reportdeffile").children().val("");
		}

	}

	function setOutputToEmail() {

		var bSelected = $("#SendToEmail").prop('checked');

		$(".reportdefemail").children().attr("readonly", !bSelected);
		button_disable($("#cmdEmailGroup")[0], !bSelected);

		if (!bSelected) {

			$(".reportdefemail").children().val("");
			$("#txtEmailGroupID").val(0);

		}

	}

	function selectEmailGroup() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtEmailGroupID").val();

		OpenHR.modalExpressionSelect("EMAIL", tableID, currentID, function (id, name) {
			$("#txtEmailGroupID").val(id);
			$("#txtEmailGroup").val(name);
		},400,400);

	}

	function selectOutputType(type) {

		$(".reportdefpreview").children().removeAttr("readonly");
		$(".reportdefscreen").children().removeAttr("readonly");
		$(".reportdeffile").children().attr("readonly", "readonly");

		switch (type) {

			case "DataOnly":
				$(".reportdefpreview").children().attr("readonly", "readonly");
				break;

			case "CSV":
				$(".reportdefscreen").children().attr("readonly", "readonly");
				$(".reportdeffile").children().removeAttr("readonly");
				break;

			case "HTML": case "WordDoc":
				$(".reportdeffile").children().removeAttr("readonly");
				$(".reportdefemail").children().removeAttr("readonly");
				break;

			case "ExcelWorksheet": case "ExcelGraph": case "ExcelPivotTable":
				$(".reportdeffile").children().removeAttr("readonly");
				$(".reportdefemail").children().removeAttr("readonly");
				break;

		}

	}

	function changeOutputType(type) {

		selectOutputType(type);

		$("#IsPreview").prop('checked', true);
		$("#ToScreen").prop('checked', true);
		$("#ToPrinter").prop('checked', false);
		$("#SaveToFile").prop('checked', false);
		$("#SendToEmail").prop('checked', false);

		setOutputToEmail();
		setOutputToFile();

		switch (type) {

			case "DataOnly":
				$("#IsPreview").prop('checked', false);
				break;

			case "CSV":
				$("#ToScreen").prop('checked', false);
				$("#SaveToFile").prop('checked', true);
				break;

			default:
				break;
		}

	}

	$(function () {
		selectOutputType('@Model.Format');

		if ('@Model.ReportType' == '@UtilityType.utlCalendarReport') {
			$(".hideforcalendarreport").hide();
		}

	});



</script>