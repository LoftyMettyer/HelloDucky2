function initialiseWidget(widgetID, containerID, widgetCaption, widgetClass) {
	main();


	function addCommas(nStr) {
		nStr += '';
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + ',' + '$2');
		}
		return x1 + x2;
	}


	/******** Our main function ********/
	function main() {
		jQuery(document).ready(function ($) {
			/******* Load CSS *******/
			var css_link = $("<link>", {
				rel: "stylesheet",
				type: "text/css",
				href: "style.css"
			});
			css_link.appendTo('head');

			//generate the HTML and bang it into the widget.

			//********** THIS IS AN OPENHR DATABASE VALUES WIDGET *****************
			// Get the data for the requested dbvalue
			// Send down the data html
			//Get the database connection object
			var widgetName = "DBValue";
			var isWidgetLogin = false;
			var widgetUser = ""; 	//"mworthing";
			var widgetPassword = ""; // "mworthing";
			var widgetDatabase = ""; // "openhr50_std";
			var widgetServer = "";	// ".\sql2012";

			$.ajax({
				url: "/dmi.net/account/getWidgetData",
				type: "POST",
				data: { widgetName: widgetName, isWidgetLogin: isWidgetLogin, widgetUser: widgetUser, widgetPassword: widgetPassword, widgetDatabase: widgetDatabase, widgetServer: widgetServer, widgetID: widgetID },
				success: function (data) {

					//clear the spinner
					$("#Spinner" + containerID).hide();

					//send the value to the browser...
					$("#" + containerID).append(document.createTextNode("   " + data.Formatting_Prefix + addCommas(data.DBValue) + "   "));




					//add marquee functionality to the relevant tiles
					if (data.DBValue > 9999) {

						$(function () {
							var scroll_text;
							$('#li_' + containerID.replace("DBV", "")).hover(
								 function () {
								 	scroll_text = setInterval(function () { scrollText(); }, 15);
								 },
								 function () {
								 	clearInterval(scroll_text);
								 	$('#' + containerID).css({
								 		left: 0
								 	});
								 }
							);

							var scrollText = function () {
								var left = $('#' + containerID).position().left - 1;
								left = -left > $('#' + containerID).width() ? $('#marquee' + containerID).width() : left;
								$('#' + containerID).css({ left: left });
							};
						});

					}
				},
				error: function (req, status, errorObj) {
					//alert("not ok"); 
					$("<p style='margin: 0px;font-size: 13pt;'>Error</p>").appendTo('#' + containerID);
				}
			});
		});
	}
}