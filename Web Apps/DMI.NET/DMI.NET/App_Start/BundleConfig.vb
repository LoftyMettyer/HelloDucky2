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
          "~/Scripts/jquery/jquery-{version}.js",
          "~/Scripts/jquery/jquery-cookie.js",
          "~/Scripts/jquery/jquery-flip.js",
          "~/Scripts/jquery/jquery-unobtrusive-ajax.js",
          "~/Scripts/jquery/jquery-validate-vsdoc.js",
          "~/Scripts/jquery/jquery-validate.js",
          "~/Scripts/jquery/jquery-validate-unobtrusive.js"))


      ' JQuery UI
      bundles.Add(New ScriptBundle("~/bundles/jQueryUI7").Include(
        "~/Scripts/jquery/jquery-{version}.js",
        "~/Scripts/jquery/jquery-ui-{version}.custom.js",
        "~/Scripts/jquery/jquery.jqGrid.src.js",
        "~/Scripts/jquery/jsTree/jquery.jstree.js",
        "~/Scripts/jquery/jquery.gridster.js",
        "~/Scripts/jquery/jquery.menu.js",
        "~/Scripts/jquery/jquery.marquee.js",
        "~/Scripts/jquery/jquery.maskedinput.js",
        "~/Scripts/jquery/jquery.mousewheel.js",
        "~/Scripts/jquery/jquery.rightClick.js",
        "~/Scripts/jquery/jquery.ui.touch-punch.min.js",
        "~/Scripts/officebar/jquery.officebar.js"))

      ' OpenHR core
      bundles.Add(New ScriptBundle("~/bundles/OpenHR_General").Include(
        "~/Scripts/openHR.js",
        "~/Scripts/FormScripts/general.js",
        "~/Scripts/FormScripts/menu.js",
        "~/Scripts/ctl_SetFont.js",
        "~/Scripts/ctl_SetStyles.js"))

      ' Custom Reports
      bundles.Add(New ScriptBundle("~/bundles/utilities_customreports").Include(
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

      ' Expression Builder
      bundles.Add(New ScriptBundle("~/bundles/utilities_expressions").Include(
        "~/Scripts/FormScripts/util_def_expression.js",
        "~/Scripts/FormScripts/util_def_exprcomponent.js"))

      ' Mail Merge
      bundles.Add(New ScriptBundle("~/bundles/utilities_mailmerge").Include(
        "~/Scripts/FormScripts/util_def_mailmerge.js"))

      ' Record Editing
      bundles.Add(New ScriptBundle("~/bundles/recordedit").Include(
        "~/Scripts/jquery/jquery.jqGrid.src.js",
        "~/Scripts/FormScripts/find.js",
        "~/Scripts/FormScripts/optionData.js",
        "~/Scripts/FormScripts/recordEdit.js"))

#If DEBUG Then

      For Each bundle As Bundle In BundleTable.Bundles
        bundle.Transforms.Clear()
      Next

#End If

    End Sub
  End Class
End Namespace