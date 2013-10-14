/*
[RC 14/10/2013]
This plugin was downloaded from http://unwrongest.com/projects/limit/ and it doesn't contain any license information
When applied to a text field it limits the number of characters that can be input;
Use it thus:
	$('#ID').limit(10);
*/

(function ($)
{
	$.fn.extend({
		limit: function (limit, element)
		{
			var interval, f;
			var self = $(this);

			$(this).focus(function ()
			{
				interval = window.setInterval(substring, 100);
			});

			$(this).blur(function ()
			{
				clearInterval(interval);
				substring();
			});

			var substringFunction = "function substring(){ var val = $(self).val();var length = val.length;if(length > limit){$(self).val($(self).val().substring(0,limit));}";
			if (typeof element != 'undefined')
				substringFunction += "if($(element).html() != limit-length){$(element).html((limit-length<=0)?'0':limit-length);}"

			substringFunction += "}";

			eval(substringFunction);
			
			substring();
		}
	});
})(jQuery);