function initialiseWidget(widgetID, containerID, widgetCaption, widgetClass) {
	main();


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
			var isWidgetLogin = true;
			var widgetUser = "mworthing";
			var widgetPassword = "mworthing";
			var widgetDatabase = "openhr50_std";
			var widgetServer = ".\sql2012";


			$("#" + containerID).addClass(widgetClass);
			$("#" + containerID).css("overflow", "hidden");
			$("#" + containerID).css("padding-top", "0px");
			$("#" + containerID).css("font-size", "60pt");
			$("#" + containerID).css("color", "white");

			$("<h1 style='font-size: 13pt'>" + widgetCaption + "</h1>").appendTo("#" + containerID);

			$.ajax({
				url: "/dmi.net/account/getWidgetData",
				type: "POST",
				data: { widgetName: widgetName, isWidgetLogin: isWidgetLogin, widgetUser: widgetUser, widgetPassword: widgetPassword, widgetDatabase: widgetDatabase, widgetServer: widgetServer, widgetID: widgetID },
				success: function (data) {
					//format the tile
					$("#" + containerID).empty();


					$("#" + containerID).addClass(widgetClass);
					$("#" + containerID).css("overflow", "hidden");
					$("#" + containerID).css("padding-top", "0px");
					$("#" + containerID).css("font-size", "60pt");
					$("#" + containerID).css("color", "white");

					//add the content
					if (data.DBValue > 999) {
						$("<p style='margin: 0px;margin-bottom: -20px;margin-top: 20px;font-size: 30pt'>" + data.DBValue + "</p>").appendTo("#" + containerID);
					}
					else {
						$("<p style='margin: 0px;margin-bottom: -20px;margin-top: 20px;font-size: 60pt'>" + data.DBValue + "</p>").appendTo("#" + containerID);
					}
					
					$("<p style='margin: 0px;font-size: 13pt;'>" + data.Formatting_Suffix + "</p>").appendTo('#' + containerID);
					$("<h1 style='font-size: 13pt'>" + data.Caption + "</h1>").appendTo("#" + containerID);

					//resize widget if number > 999
					//if (data.DBValue > 999) { resizeWidget(containerID, 2, 1); }
				},
				error: function (req, status, errorObj) {
					//alert("not ok"); 
					$("#" + containerID).empty();
					$("<p style='margin: 0px;font-size: 13pt;'>Error</p>").appendTo('#' + containerID);
				}
			});
		});
	}
}