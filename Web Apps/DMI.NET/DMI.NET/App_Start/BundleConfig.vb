Imports System.Web.Optimization

Namespace App_Start

	Public Class BundleConfig

		Public Shared Sub RegisterBundles(bundles As BundleCollection)

			' Microsoft elements
			Dim codeBundle As New ScriptBundle("~/bundles/Microsoft")
			codeBundle.IncludeDirectory("~/Scripts/Microsoft", "*.js")
			bundles.Add(codeBundle)


			' JQuery core
			bundles.Add(New ScriptBundle("~/bundles/jQuery").Include(
					"~/Scripts/jquery-{version}.js",
					"~/Scripts/jquery/jquery.cookie.js",
					"~/Scripts/jquery/jquery-flip.js",
					"~/Scripts/jquery/jquery-unobtrusive-ajax.js",
					"~/Scripts/jquery-validate-vsdoc.js",
					"~/Scripts/jquery-validate.js",
					"~/Scripts/jquery-validate-unobtrusive.js",
					"~/Scripts/date.js"))


			' JQuery UI
			bundles.Add(New ScriptBundle("~/bundles/jQueryUI7").Include(
				"~/Scripts/jquery-{version}.js",
				"~/Scripts/jquery-ui-{version}.js",
				"~/Scripts/jquery.jqGrid.js",
				"~/Scripts/jquery/grid.locale-en.js",
				"~/Scripts/jquery/jsTree/jquery.jstree.js",
				"~/Scripts/jquery/jquery.gridster.js",
				"~/Scripts/jquery/jquery.menu.js",
				"~/Scripts/jquery/jquery.marquee.js",
				"~/Scripts/jquery/jquery.maskedinput.js",
				"~/Scripts/jquery/jquery.mousewheel.js",
				"~/Scripts/jquery/jquery.numeric.js",
				"~/Scripts/jquery/jquery.rightClick.js",
				"~/Scripts/jquery/jquery.ui.touch-punch.min.js",
				"~/Scripts/jquery/jOrgChart/prettify.js",
				"~/Scripts/jquery/jOrgChart/jquery.jOrgChart.js",
				"~/Scripts/officebar/jquery.officebar_MODIFIED.js"))

			' OpenHR core
			bundles.Add(New ScriptBundle("~/bundles/OpenHR_General").Include(
				"~/Scripts/openHR.js",
				"~/Scripts/clientDLL.js",
				"~/Scripts/FormScripts/general.js",
				"~/Scripts/FormScripts/menu.js",
				"~/Scripts/ctl_SetStyles.js"))

			' Custom Reports
			bundles.Add(New ScriptBundle("~/bundles/utilities_customreports").Include(
				"~/Scripts/FormScripts/ReportDefinition.js",
				"~/Scripts/FormScripts/customreport.js",
				"~/Scripts/FormScripts/util_def_customreports.js",
				"~/Scripts/FormScripts/general.js"))

			' Cross Tabs
			bundles.Add(New ScriptBundle("~/bundles/utilities_crosstabs").Include(
				"~/Scripts/FormScripts/crosstab.js",
				"~/Scripts/FormScripts/crosstabdef.js",
				"~/Scripts/FormScripts/general.js"))

			' Calendar Reports
			bundles.Add(New ScriptBundle("~/bundles/utilities_calendarreports").Include(
				"~/Scripts/FormScripts/calendarreportdef.js",
				"~/Scripts/FormScripts/general.js"))

			' Standard Reports
			bundles.Add(New ScriptBundle("~/bundles/utilities_standardreports").Include(
				"~/Scripts/FormScripts/stdrpt_def_absence.js"))

			' Expression Builder
			bundles.Add(New ScriptBundle("~/bundles/utilities_expressions").Include(
				"~/Scripts/FormScripts/util_def_expression.js",
				"~/Scripts/FormScripts/util_def_exprcomponent.js"))

			' Mail Merge
			bundles.Add(New ScriptBundle("~/bundles/utilities_mailmerge").Include(
				"~/Scripts/FormScripts/util_def_mailmerge.js"))

			' Picklists
			bundles.Add(New ScriptBundle("~/bundles/utilities_picklists").Include(
				"~/Scripts/FormScripts/util_def_picklist.js"))

			' Record Editing
			bundles.Add(New ScriptBundle("~/bundles/recordedit").Include(
				"~/Scripts/jquery/jquery.jqGrid.src.js",
				"~/Scripts/FormScripts/find.js",
				"~/Scripts/FormScripts/optionData.js",
				"~/Scripts/FormScripts/recordEdit.js",
				"~/Scripts/autoNumeric-1.9.15.js",
				"~/Scripts/ColorPicker/spectrum.js"))

			'ActiveX bundle (no minify)
			bundles.Add(New Bundle("~/bundles/OpenHR_ActiveX").Include(
				"~/Scripts/ctl_SetFont.js"))

			' OptionData grid bundle
			bundles.Add(New ScriptBundle("~/bundles/optiondatagrid").Include(
										 "~/Scripts/FormScripts/OptiondataGrid.js"))

#If DEBUG Then

			For Each bundle As Bundle In BundleTable.Bundles
				bundle.Transforms.Clear()
			Next

#End If

		End Sub
	End Class
End Namespace