
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of ReportColumnItem)

@If (Model.DataType = ColumnDataType.sqlOle Or Model.DataType = ColumnDataType.sqlVarBinary) Then
@<div style="text-align: center;padding:4px">
   <img src="@Model.ColumnValue" style="height:@Model.Height.ToString()px;">
</div>
Else
   Dim className As String

   If (Model.DefaultHeight = 1) Then
      className = "truncate"
   Else
      className = "divMultiline"
   End If
   @<div class="@className" style="min-height:20px;height:@Model.Height.ToString()px;text-align: center;margin:2px;">
      @If Model.ColumnValue = String.Empty Then
      @<span style="font-size:@Model.FontSize.ToString()px;">&nbsp;</span>
      Else
      @<span title="@Model.ColumnTitle" style="font-size:@Model.FontSize.ToString()px;">@Model.ColumnValue</span>
      End If
   </div>
End If