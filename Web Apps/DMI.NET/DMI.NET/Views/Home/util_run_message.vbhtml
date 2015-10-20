@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Imports DMI.NET.Code.Extensions
@Inherits WebViewPage(Of Responses.PostResponse)

<input type='hidden' id="txtDefn_ErrMsg" name="txtDefn_ErrMsg" value="@Model.Message.Replace(vbNewLine, "<br/>")">

<script type="text/javascript">
    OpenHR.modalPrompt($("#txtDefn_ErrMsg").val(), 2, "", "");
    closeclick();
</script>