@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.Home.BulkBookingSelectionViewModel)
@Imports DMI.NET
@imports system.web.optimization

<div id="BulkBookingSelection" class="absolutefull optiondatagridpage">
	<div class="pageTitleDiv">
		<span class="pageTitle">Select Records</span>
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
		<div class="clearboth" id="FindGridRow">
			<table id="ssOleDBGridSelRecords" style="width: 100%"></table>
			<div id="ssOLEDBPager" style=""></div>
		</div>
	</main>	
</div>

@Html.HiddenFor(Function(m) m.TableID, New With {.id = "txtTableID"})
@Html.HiddenFor(Function(m) m.PageAction, New With {.id = "txtPageAction"})

@Scripts.Render("~/bundles/bulkbookingselection")