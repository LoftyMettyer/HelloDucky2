function initialiseWidget(widgetID, containerID, widgetCaption, widgetClass) {


	/******* Load CSS *******/
	var css_link = $("<link>", {
		rel: "stylesheet",
		type: "text/css",
		href: window.ROOT + "Content/widgetCSS/wdg_oHRHoliday.css"
	});
	css_link.appendTo('head');

	var tmpCssText = "";

	if (widgetCaption == null || widgetCaption <= 0) widgetCaption = 1;
	
	var PctComplete = (widgetCaption / 25) * 100;

	var bulletbox = document.createElement('div');
	bulletbox.className = "bulletbox";

	var actual = document.createElement('div');
	actual.className = "actual";

	var uislider = document.createElement('div');
	uislider.className = "uislider";
	uislider.style.cssText = 'width: ' + PctComplete + '%;';	

	var uisliderhandle = document.createElement('a');
	uisliderhandle.className = "uisliderhandle";
	uisliderhandle.style.cssText = 'left: ' + PctComplete + '%;';

	var markrange0 = document.createElement('div');
	markrange0.className = "markrange0";

	var markrange1 = document.createElement('div');
	markrange1.className = "markrange1";	
   markrange1.style.cssText = 'left: ' + PctComplete + '%;';

   var captionrange0 = document.createElement('div');
	captionrange0.className = "captionrange0";
	captionrange0.style.cssText = 'width: ' + PctComplete + '%;';
	
	for (var i = 0; i < 101; i+=20 ) {
		this.markBottom = {};
		this.captionBottom = {};
		
		var markBottom = "markbottom" + i;
		var captionBottom = "captionbottom" + i;

		this[markBottom] = document.createElement('div');
		this[markBottom].className = "markBottom";
		this[captionBottom] = document.createElement('div');
		this[captionBottom].className = "captionBottom";

		
		tmpCssText += 'left: ' + i + '%; ';
		this[markBottom].style.cssText = tmpCssText;
		
		tmpCssText = 'left: ' + i + '%; ';
		this[captionBottom].style.cssText = tmpCssText;
		this[captionBottom].textContent = i / 4;

		bulletbox.appendChild(this[markBottom]);
		bulletbox.appendChild(this[captionBottom]);
	}

	var info = document.createElement('div');
	info.className = "info";	
	info.textContent = widgetCaption;

	//build content	
	actual.appendChild(uislider);
	actual.appendChild(uisliderhandle);

	captionrange0.textContent = "Holiday Taken";

	bulletbox.appendChild(markrange0);
	bulletbox.appendChild(markrange1);
	bulletbox.appendChild(captionrange0);

	bulletbox.appendChild(actual);
	
	//adjust parent container
	//see css
	
	//add content
	$("#" + containerID).append(bulletbox);
	$("#" + containerID).append(info);

}