function initialiseWidget(widgetID, containerID, widgetCaption, widgetClass) {

		//format the tile
		$("#" + containerID).empty();
		$("#" + containerID).addClass(widgetClass);
		$("#" + containerID).css("overflow", "hidden");
		$("#" + containerID).css("padding-top", "45px");
		$("#" + containerID).css("font-size", "60pt");
		$("#" + containerID).css("color", "white");
					
		$("<h1 style='font-size: 13pt;height: 2.5em;'>" + widgetCaption + "</h1>").appendTo("#" + containerID);

}