@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.TalentReportModel)


<fieldset id="MatchTables" class="floatleft overflowhidden width50">
	<legend class="fontsmalltitle">Data :</legend>

	<fieldset class="">
		Role Match Table :
		<select class="width70 floatright enableSaveButtonOnComboChange" name="BaseChildTableID" id="BaseChildTableID" onchange="refreshTalentReportBaseColumns(event.target);"></select>
	</fieldset>

	<fieldset>
		Match Column : <select class="width70 floatright enableSaveButtonOnComboChange" name="BaseChildColumnID" id="BaseChildColumnID" onchange="SetSelectedColumnDataType('Base')"></select>
	</fieldset>

	<fieldset>
		Minimum Rating : <select class="width70 floatright enableSaveButtonOnComboChange" name="BaseMinimumRatingColumnID" id="BaseMinimumRatingColumnID"></select>
	</fieldset>

	<fieldset>
		Preferred Rating : <select class="width70 floatright enableSaveButtonOnComboChange" name="BasePreferredRatingColumnID" id="BasePreferredRatingColumnID"></select>
	</fieldset>

	<br />
	<fieldset class="">
		Person Match Table : <select class="width70 floatright enableSaveButtonOnComboChange" name="MatchChildTableID" id="MatchChildTableID" onchange="refreshTalentReportMatchColumns(event.target);"></select>
	</fieldset>
	<fieldset class="">
		Match Column : <select class="width70 floatright enableSaveButtonOnComboChange" name="Matchchildcolumnid" id="MatchChildColumnID" onchange="SetSelectedColumnDataType('Match')"></select>
	</fieldset>
	<fieldset class="">
		Actual Rating : <select class="width70 floatright enableSaveButtonOnComboChange" name="MatchChildRatingColumnID" id="MatchChildRatingColumnID"></select>
	</fieldset>

  </fieldset>

  <fieldset id="MatchTables" class="floatleft overflowhidden width50">
    <legend class="fontsmalltitle">Match Filter :</legend>

    <fieldset>
      Match Against :
      @Html.RadioButton("matchagainsttype", MatchAgainstType.Any, Model.MatchAgainstType = MatchAgainstType.Any, New With {.id = "matchagainsttype_any"})
      Any
      @Html.RadioButton("matchagainsttype", MatchAgainstType.All, Model.MatchAgainstType = MatchAgainstType.All, New With {.id = "matchagainsttype_all"})
      All
    </fieldset>

    <fieldset>
      @Html.LabelFor(Function(m) m.MinimumScore)
      @Html.TextBoxFor(Function(m) m.MinimumScore)
    </fieldset>

    <fieldset>
      @Html.CheckBoxFor(Function(m) m.IncludeUnmatched)
      @Html.LabelFor(Function(m) m.IncludeUnmatched)
    </fieldset>
		@Html.HiddenFor(Function(m) m.BaseChildColumnDataType, New With {.id = "BaseChildColumnDataType"})
		@Html.HiddenFor(Function(m) m.MatchChildColumnDataType, New With {.id = "MatchChildColumnDataType"})
  		@Html.HiddenFor(Function(m) m.MatchChildTableName, New With {.id = "MatchChildTableName"})
  		@Html.HiddenFor(Function(m) m.BaseChildTableName, New With {.id = "BaseChildTableName"})
  		@Html.HiddenFor(Function(m) m.BaseTableName, New With {.id = "BaseTableName"})
  		@Html.HiddenFor(Function(m) m.MatchTableName, New With {.id = "MatchTableName"})
  </fieldset>

