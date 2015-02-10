@Imports DMI.NET
@Imports System.Web.Optimization
@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.Account.ConfigurationErrorsModel)

@Styles.Render("~/bundles/stylesheets")

<div class="verticalpadding200"></div>

<div class="ui-widget-content ui-corner-tl ui-corner-br loginframe">
	<img class="loginframeImage" alt="loginimage" src="@Url.Content("~/Content/images/systemerror.jpg")">
	<div class="aligncenter">
		<h3>
			@For Each item In Model.Errors
				@item.Detail
			Next
		</h3>
	</div>
</div>

<br/>

<p class="aligncenter">
	Please contact your system administrator.
</p>


