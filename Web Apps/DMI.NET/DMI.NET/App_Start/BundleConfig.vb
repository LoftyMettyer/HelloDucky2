Imports System.Web.Optimization

Public Class BundleConfig

  Public Shared Sub RegisterBundles(bundles As BundleCollection)

    bundles.Add(New ScriptBundle("~/bundles/jquery").Include(
                "~/Scripts/jquery-{version}.js"))

    bundles.Add(New ScriptBundle("~/bundles/jqueryui").Include(
                "~/Scripts/jquery-ui-{version}.js",
                "~/Scripts/jquery.gridster.js"))


  End Sub
End Class
