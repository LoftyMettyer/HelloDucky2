﻿
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of ReportColumnItem)

@If (Model.DataType = ColumnDataType.sqlOle Or Model.DataType = ColumnDataType.sqlVarBinary) Then
@<div style="text-align: center;padding:4px">
   <img src="@Model.ColumnValue" style="height:@Model.Height.ToString()px;">
</div>  
Else
   If (Model.DefaultHeight = 1) Then
      @<div class="truncate" style="min-height:20px;height:@Model.Height.ToString()px;text-align: center;margin:2px;">
         <span title="@Model.ColumnValue" style="font-size:@Model.FontSize.ToString()px;">@Model.ColumnValue</span>
      </div>  
   Else
      @<div class="divMultiline" style="min-height:20px;height:@Model.Height.ToString()px;text-align: center;margin:2px;">
         <span title="@Model.ColumnValue" style="font-size:@Model.FontSize.ToString()px;">@Model.ColumnValue</span>
      </div>
   End If
End If