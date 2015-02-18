@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.Home.BulkBookingViewModel)
@Imports DMI.NET
@imports system.web.optimization

@Using (Html.BeginForm("BulkBooking_Submit", "Home", FormMethod.Post, New With {.id = "frmBulkBooking", .name = "frmBulkBooking", .defaultbutton = "cmdOK"}))
	@<div class="absolutefull optiondatagridpage">
	<div class="ML20px">
		<div class="pageTitleDiv">
			<span class="pageTitle">Bulk Booking</span>
		</div>
		<nav>
			@If Model.TbStatusPExists = True Then
				@<div class="formField floatleft">
					@Html.LabelFor(Function(m) m.BookingStatus)
					@Html.DropDownListFor(Function(m) m.BookingStatuses, New SelectList(Model.BookingStatuses, "Value", "Text", Model.BookingStatuses.First().Value), New With {.id = "selStatus"})
					@Html.ValidationMessageFor(Function(m) m.BookingStatuses)
				</div>
			End If
		</nav>
		<main class="clearboth">
			<div class="stretchyfill" id="FindGridRow" style="height: 400px; margin-bottom: 50px;">
				<table id="ssOleDBGridFindRecords" style="width: 100%"></table>
			</div>
			<div class="navButtons stretchyfixed">
				<button type="button" id="cmd_tbBBSelect">Add</button>
				<button type="button" id="cmd_tbBBFilteredAdd">Filtered Add</button>
				<button type="button" id="cmd_tbBBPicklistAdd">Picklist Add</button>
				<button type="button" id="cmd_tbBBRemove">Remove</button>
				<button type="button" id="cmd_tbBBRemoveAll">Remove All</button>
				<button type="button" id="cmd_tbBBOK">OK</button>
				<button type="button" id="cmd_tbBBCancel">Cancel</button>
			</div>
		</main>
	</div>
	@Html.AntiForgeryToken()
</div>
End Using

@Using (Html.BeginForm("BulkBooking_Submit", "Home", FormMethod.Post, New With {.id = "frmGotoOption", .name = "frmGotoOption", .defaultbutton = "cmdOK"}))
	Html.RenderPartial("~/Views/Shared/gotoOption.ascx")
	@Html.AntiForgeryToken()
End Using

@Html.HiddenFor(Function(m) m.TableID)
@Html.HiddenFor(Function(m) m.txt1000SepCols)
@Html.HiddenFor(Function(m) m.TbStatusPExists)
@Html.HiddenFor(Function(m) m.CourseRecordID, New With {.id = "txtOptionRecordID"})

@Scripts.Render("~/bundles/bulkbooking")
