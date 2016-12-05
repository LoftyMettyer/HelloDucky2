@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Imports System.Linq
@Inherits System.Web.Mvc.WebViewPage(Of OrganisationReportModel)
<style>
   .truncate {
      width: 95%;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
   }
</style>

<div style="width:300px;border:1px solid gray;padding:5px;overflow:auto;max-height:300px" class="divMainContainer centered">
   @If (Model.PostBasedTableId > 0) Then
      For Each item In Model.PreviewColumnList.Where(Function(m) m.ViewID = Model.BaseViewID)
         Html.RenderPartial("_PreviewOrganisationColumn", item)
      Next
      @If (Model.PreviewColumnList.Where(Function(m) m.ViewID <> Model.BaseViewID).Count > 0) Then
      @<div Style="margin:10px;border:1px solid gray;padding:5px;" Class="centered">
         @For Each itemchild In Model.PreviewColumnList.Where(Function(m) m.ViewID <> Model.BaseViewID)
            Html.RenderPartial("_PreviewOrganisationColumn", itemchild)
         Next
      </div>
      End If
   Else
      For Each item In Model.PreviewColumnList
         Html.RenderPartial("_PreviewOrganisationColumn", item)
      Next
   End If

</div>
<br /><br />
<table class="centered">
   <tr>
      <td>
         <div id="divButtons" Class="clearboth">
            <input type="button" value="OK" id="butEditChildTableCancel" onclick="closePopup();" />
         </div>
      </td>
   </tr>
</table>

<script type="text/javascript">

    $(document).ready(function () {
        setTimeout(function() {
            $(".divMultiline").dotdotdot({ wrap: 'letter', fallbackToLetter: true });
        }, 1);
    });

    function closePopup() {
        $("#divPopupPreview").dialog("close");
        $("#divPopupPreview").empty();
    }
</script>
