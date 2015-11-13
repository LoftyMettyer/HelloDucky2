@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.MailMergeModel)
@Html.HiddenFor(Function(m) m.ActionType)

<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />

<div>
    <fieldset>
        <legend class="fontsmalltitle">Template:</legend>
        <div class="floatleft width100">
            <div class="upload">
                @Html.Hidden("txtMaxRequestLength", Session("maxRequestLength"))
                @Html.TextBoxFor(Function(m) m.UploadTemplateName, New With {.id = "txtTemplateFileName", .class = "width25", .ReadOnly = "ReadOnly"})
                <label for="TemplateFile">Upload</label>
                <input id="button_download_template" type="button" value="Download" onclick="DownloadTemplate();" />
            </div>
        </div>
    </fieldset>
    <br />

    <fieldset class="floatleft">
        <legend class="fontsmalltitle">Options:</legend>
        @Html.CheckBoxFor(Function(m) m.PauseBeforeMerge)
        @Html.LabelFor(Function(m) m.PauseBeforeMerge)
        <br />
        @Html.CheckBoxFor(Function(m) m.SuppressBlankLines)
        @Html.LabelFor(Function(m) m.SuppressBlankLines)
        <br /><br /><br />
    </fieldset>
</div>

<fieldset class="width25 floatleft clearboth" style="">
	<legend class="fontsmalltitle">Output Format:</legend>
	<fieldset class="">
		<div class="margebot10">
			@Html.RadioButton("OutputFormat", 0, Model.OutputFormat = MailMergeOutputTypes.WordDocument, New With {.onclick = "selectMergeOutputType('WordDocument')"})
                    Word Document
			<br />
		</div>
        <div class="margebot10">
			@Html.RadioButton("OutputFormat", 1, Model.OutputFormat = MailMergeOutputTypes.IndividualEmail, New With {.onclick = "selectMergeOutputType('IndividualEmail')"})
			Individual Emails
                    <br />
		</div>
		@Html.RadioButton("OutputFormat", 2, Model.OutputFormat = MailMergeOutputTypes.DocumentManagement, New With {.onclick = "selectMergeOutputType('DocumentManagement')"})
		<span class="DataManagerOnly">Document Management</span>
	</fieldset>
</fieldset>

<fieldset class="outputmerge_WordDocument width60 floatleft" style="">
	<legend class="fontsmalltitle">Word Document:</legend>
	<fieldset>
        <div class="padbot10">
			@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen, New With {.id = "WordDisplayOutputOnScreen"})
			@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
			<br />
		</div>

		<div class="reportdefprinter DataManagerOnly">
			<div class="width30 floatleft margebot10">
				@Html.CheckBoxFor(Function(m) m.SendToPrinter, New With {.id = "SendToPrinter", .onclick = "setSendToPrinter();"})
				@Html.LabelFor(Function(m) m.SendToPrinter)
			</div>
			<div class="width70 floatleft padbot5 margebot10">
				@Html.TextBoxFor(Function(m) m.WordDocumentPrinter, New With {.placeholder = "Default printer", .class = "DataManagerOnly readonly width100"})
			</div>
		</div>

		<div class="padbot5">
			<div class="width30 floatleft">
				@Html.CheckBoxFor(Function(m) m.SaveToFile, New With {.id = "SaveToFile", .onclick = "setOutputToFile();"})
				@Html.LabelFor(Function(m) m.SaveToFile)
			</div>
			<div class="width70 floatleft">
				@Html.TextBoxFor(Function(m) m.Filename, New With {.id = "Filename", .placeholder = "File Name", .class = "outputfile width100"})
				@Html.ValidationMessageFor(Function(m) m.Filename)
			</div>
		</div>
	</fieldset>
</fieldset>

<fieldset class="outputmerge_IndividualEmail width60 floatleft" style="">
	<legend class="fontsmalltitle">Individual Emails:</legend>

	<fieldset id="fieldsetsubjectemail">
		@Html.LabelFor(Function(m) m.EmailGroupID, New With {.class = "display-label_emails"})
		@Html.EmailGroupDropdown("EmailGroupID", Model.EmailGroupID, Model.AvailableEmails)
		<br />
		@Html.LabelFor(Function(m) m.EmailSubject, New With {.class = "display-label_emails"})
		@Html.TextBox("EmailSubject", Model.EmailSubject, New With {.class = "display-textbox-emails"})
		<br />
		<br />
		@Html.CheckBoxFor(Function(m) m.EmailAsAttachment, New With {.id = "EmailAsAttachment", .onclick = "setOutputSendAsAttachment();"})
		@Html.LabelFor(Function(m) m.EmailAsAttachment)
		<br />
		@Html.LabelFor(Function(m) m.EmailAttachmentName, New With {.class = "display-label_emails margeTop10"})
		@Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {.id = "EmailAttachmentName", .Name = "EmailAttachmentName", .class = "display-textbox-emails  margeTop10"})
		<br />
		@Html.ValidationMessageFor(Function(m) m.EmailAttachmentName)
		<br />
		@Html.ValidationMessageFor(Function(m) m.EmailGroupID)
	</fieldset>

</fieldset>

<fieldset class="outputmerge_DocumentManagement width60 floatleft" style="">
	<legend class="fontsmalltitle">	Document Management :</legend>
	<fieldset>
        <div class="padbot5">
			<div class="width30 floatleft">
				@Html.LabelFor(Function(m) m.PrinterName)
			</div>
			<div class="width70 floatleft padbot5">			
				@Html.TextBoxFor(Function(m) m.DocumentManagementPrinter, New With {.placeholder = "Default printer", .class = "DataManagerOnly readonly width100"})			
			</div>
		</div>
		<br />		
		<br />		
		<div class="padbot5">				
			@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen, New With {.id = "DocumentDisplayOutputOnScreen"})
			@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)			
		</div>

	</fieldset>
</fieldset>

<fieldset class="DataManagerOnly width100">
	Note: Options marked in red are unavailable in OpenHR Web.
</fieldset>

<script type="text/javascript">

    function SubmitTemplate() {

        var maxRequestLength = Number($("#txtMaxRequestLength").val());
        maxRequestLength = Math.min(maxRequestLength, (4096 * 1024));

        var lngFileSize = $("#TemplateFile")[0].files[0].size;

        if (lngFileSize > maxRequestLength) {
            OpenHR.modalMessage("Template is too large to upload. \nMaximum file upload size for this is " + (maxRequestLength / 1024) + "KB", 48);
            return false;
        }

        var filename = $("#TemplateFile").val().replace(/^.*[\\\/]/, '');
        $("#txtTemplateFileName").val(filename);

        var form = document.getElementById("frmTemplateFile");
        var data = new FormData(form);

        $.ajax({
            type: "POST",
            url: "reports\\util_def_mailmerge_submittemplate",
            contentType: false,
            processData: false,
            data: data,
            success: function (result) {
                button_disable($("#button_download_template")[0], false);
                OpenHR.modalMessage("Template uploaded successfully");
            },
            error: function (xhr, status, p3, p4) {
                var err = p3;
                if (xhr.responseText && xhr.responseText[0] == "{")
                    err = JSON.parse(xhr.responseText).Message;
                $("#txtTemplateFileName").val("");
                OpenHR.modalMessage(err);
            }
        });


        menu_toolbarEnableItem('mnutoolSaveReport', true);
    }

    function DownloadTemplate() {
        var frmDownloadTemplate = $("#frmDownloadTemplate")[0];
        frmDownloadTemplate.submit();
    }

    function setOutputToFile() {
		var bSelected = $("#SaveToFile").prop("checked");
		$(".outputfile").children().attr("readonly", !bSelected);
		if (!bSelected) {
			$(".outputfile").children().val("");
			$('#Filename').val("");
		}
		$('#Filename').prop('disabled', !bSelected);
	}

    function setOutputSendAsAttachment() {
		var bSelected = $("#EmailAsAttachment").prop("checked");
		$("#EmailAttachmentName").attr("readonly", !bSelected);

		if (!bSelected) {
			$("#EmailAttachmentName").val("");
		}
	}

    function selectMergeOutput(outputType) {
		$("[class^=outputmerge_]").hide();
		$(".outputmerge_" + outputType).show(500);
	}

    function selectMergeOutputType(outputType) {		
		if (outputType == 'DocumentManagement') {
			$('#DocumentDisplayOutputOnScreen').prop('checked', true);
			$('#WordDocumentPrinter').prop('disabled', true);
			$('#SendToPrinter').prop('checked', false);
		}
		else if (outputType == 'WordDocument') {
			$('#WordDisplayOutputOnScreen').prop('checked', true);
			$('#SendToPrinter').prop('checked', false);
			$('#SaveToFile').prop('checked', false);
			$('#Filename').val("");
			$('#Filename').prop('disabled', true);
		}
		else if (outputType == 'IndividualEmail') {
			$('#EmailAsAttachment').prop('checked', false);
			$('#EmailSubject').val("");
			$('#EmailAttachmentName').val("");
		}

		$("[class^=outputmerge_]").hide();
		$(".outputmerge_" + outputType).show(500);
	}

    function setSendToPrinter()
	{
		var bSelected = $("#SendToPrinter").prop("checked");
		if (bSelected) {
			$('#WordDocumentPrinter').prop('disabled', false);
		}
		else
		{
			$('#WordDocumentPrinter').prop('disabled', true);
		}
	}

	$(function () {

		selectMergeOutput('@Model.OutputFormat');
        setOutputSendAsAttachment();
		//styling for email address under Individual Emails section
		$('#fieldsetsubjectemail select').css({
			width: '69%',
marginBottom: '10px',
float: 'left'
		});

		$('#EmailSubject, #EmailAttachmentName').css({
			width: '68.3%',
float: 'left'
		});		

		$('fieldset').css("border", "1");
		$('#WordDisplayOutputOnScreen').prop('checked', true);
		$('#WordDocumentPrinter').prop('disabled', true);
		if ('@Model.Filename' == '')
        {
			$('#Filename').prop('disabled', true);
		}
        if ($("#ActionType").val() == '@UtilityActionType.New') {
			$('#PauseBeforeMerge').prop('checked', true);
			$('#SuppressBlankLines').prop('checked', true);
		}

	    $('.upload label').button();

	    if ($("#txtTemplateFileName").val() === "") {
	        button_disable($("#button_download_template")[0], true);
	    }

	});

</script>