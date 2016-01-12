@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.TalentReportModel)


<fieldset id="MatchTables" class="floatleft overflowhidden width50">
	<legend class="fontsmalltitle">Data :</legend>

	<fieldset class="">
		Role Match Table :
		<select class=" width70 floatright" name=" basechildtableid" id="BaseChildTableID" onchange="refreshTalentReportBaseColumns(event.target);"></select>
	</fieldset>

	<fieldset>
		Match Column : <select class="width70 floatright" name="BaseChildColumnID" id="BaseChildColumnID"></select>
	</fieldset>

	<fieldset>
		Minimum Rating : <select class="width70 floatright" name="BaseMinimumRatingColumnID" id="BaseMinimumRatingColumnID"></select>
	</fieldset>

	<fieldset>
		Preferred Rating : <select class="width70 floatright" name="BasePreferredRatingColumnID" id="BasePreferredRatingColumnID"></select>
	</fieldset>

	<br />
	<fieldset class="">
		Person Match Child : <select class="width70 floatright" name="MatchChildTableID" id="MatchChildTableID" onchange="refreshTalentReportMatchColumns(event.target);"></select>
	</fieldset>
	<fieldset class="">
		Match Column : <select class="width70 floatright" name="MatchChildColumnID" id="MatchChildColumnID"></select>
	</fieldset>
	<fieldset class="">
		Actual Rating : <select class="width70 floatright" name="MatchChildRatingColumnID" id="MatchChildRatingColumnID"></select>
	</fieldset>

	<br />
	<div class="width70 floatright">
		@Html.RadioButton("matchagainsttype", MatchAgainstType.Any, Model.MatchAgainstType = MatchAgainstType.Any, New With {.id = "matchagainsttype_any"})
		Match Against Any
		@Html.RadioButton("matchagainsttype", MatchAgainstType.All, Model.MatchAgainstType = MatchAgainstType.All, New With {.id = "matchagainsttype_all"})
		Match Against All
	</div>


</fieldset>
