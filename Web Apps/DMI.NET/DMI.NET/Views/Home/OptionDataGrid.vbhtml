﻿@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.OptionDataGridViewModel)
@Imports DMI.NET
@imports system.web.optimization

<div class="absolutefull optiondatagridpage">
	<div class="pageTitleDiv">
		<span class="pageTitle">@Model.PageTitle</span>
	</div>
	<nav>
		<div class="formField floatleft">
			@Html.LabelFor(Function(m) m.View)
			@Html.DropDownListFor(Function(m) m.Views, New SelectList(Model.Views, "Value", "Text", Model.Views.First().Value), New With {.id = "selectView"})
			@Html.ValidationMessageFor(Function(m) m.Views)
		</div>
		<div class="formField floatright">
			@Html.LabelFor(Function(m) m.Order)
			@Html.DropDownListFor(Function(m) m.Orders, New SelectList(Model.Orders, "Value", "Text", Model.Orders.First().Value), New With {.id = "selectOrder"})
			@Html.ValidationMessageFor(Function(m) m.Orders)
		</div>
	</nav>
	<main>
		<div class="clearboth" id="FindGridRow" style="margin-bottom: 50px;">
			<table id="ssOleDBGridRecords" style="width: 100%"></table>
			<div id="pager-coldata-optiondata"></div>
		</div>
	</main>
	<footer>
		<button id="cmdSelect">Select</button>
		<button id="cmdCancel">Cancel</button>
	</footer>
</div>

@Using (Html.BeginForm(Model.SubmitAction, "Home", FormMethod.Post,
											New With {.id = "frmGotoOption", .name = "frmGotoOption", .defaultbutton = "cmdSelect"}))
	Html.RenderPartial("~/Views/Shared/gotoOption.ascx")
End Using

@Html.HiddenFor(Function(m) m.RecordID)
@Html.HiddenFor(Function(m) m.TableID)
@Html.HiddenFor(Function(m) m.CourseTitle)

@Html.HiddenFor(Function(m) m.DataFrameSource)
@Html.HiddenFor(Function(m) m.OptionAction)
@Html.HiddenFor(Function(m) m.GotoOptionActionSelect)
@Html.HiddenFor(Function(m) m.GotoOptionActionCancel)

@Scripts.Render("~/bundles/optiondatagrid")