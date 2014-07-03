@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums

@Inherits System.Web.Mvc.WebViewPage(Of Models.MailMergeModel)

@Code
	Layout = Nothing
End Code

<style>
	.mergeoutput {
		/*display: none;*/
		width: 50%;
		float: left;
	}
	
</style>


<div>
	Options:
	<br/>

	@Html.LabelFor(Function(m) m.TemplateFileName)
	@Html.TextBox("TemplateFileName", Model.TemplateFileName)
	<input type="button" class="ui-state-disabled" id="cmdEmailGroup" name="cmdTemplate" value="..." style="padding-top: 0;" />

	@Html.CheckBoxFor(Function(m) m.PauseBeforeMerge)
	@Html.LabelFor(Function(m) m.PauseBeforeMerge)
	<br/>

	@Html.CheckBoxFor(Function(m) m.SuppressBlankLines)
	@Html.LabelFor(Function(m) m.SuppressBlankLines)

</div>

<br />

<div>
	<div class="left">

		Output Format:
		<br/>

		@Html.RadioButton("OutputFormat", 0, Model.OutputFormat = MailMergeOutputTypes.WordDocument, New With {.onclick = "selectMergeOutput(0)"})
		Word Document
		<br />

		@Html.RadioButton("OutputFormat", 1, Model.OutputFormat = MailMergeOutputTypes.IndividualEmail, New With {.onclick = "selectMergeOutput(1)"})
		Individual Emails
		<br />

		@Html.RadioButton("OutputFormat", 2, Model.OutputFormat = MailMergeOutputTypes.DocumentManagement, New With {.onclick = "selectMergeOutput(2)"})
		Document Management

	</div>
	<div class="mergeoutput"  id="divMergeOutput_0">
		Word Document: 
		<br/>
		@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
		@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
		<br/>

		@Html.CheckBoxFor(Function(m) m.SendToPrinter)
		@Html.LabelFor(Function(m) m.SendToPrinter)
		Printer location:	@Html.TextBox("PrinterName", Model.PrinterName)
		<br />

		@Html.CheckBoxFor(Function(m) m.SaveTofile)
		@Html.LabelFor(Function(m) m.SaveTofile)
		@Html.TextBox("Filename", Model.Filename)
		<br />

	</div>

	<br />
	<div class="mergeoutput" id="divMergeOutput_1">
		Individual Emails:
		<br />
		@Html.LabelFor(Function(m) m.EmailGroupID)
		@Html.TextBox("EmailGroupID", Model.EmailGroupID)
		<br />
		@Html.LabelFor(Function(m) m.EmailSubject)
		@Html.TextBox("EmailSubject", Model.EmailSubject)

		<br />

		@Html.LabelFor(Function(m) m.EmailAsAttachment)
		@Html.CheckBoxFor(Function(m) m.EmailAsAttachment)
		<br />

		@Html.LabelFor(Function(m) m.EmailAttachmentName)
		@Html.TextBox("EmailAttachmentName", Model.EmailAttachmentName)

	</div>

	<div class="mergeoutput" id="divMergeOutput_2">
		Document Management :
		<br />
		@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
		@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)

		<br />
		Engine:		@Html.TextBox("PrinterName", Model.PrinterName)
		<br />

	</div>


</div>

<script type="text/javascript">

	function selectMergeOutput(outputType) {

		$("#divMergeOutput_").hide();
		$("#divMergeOutput_" + outputType).show(500);

	}

</script>