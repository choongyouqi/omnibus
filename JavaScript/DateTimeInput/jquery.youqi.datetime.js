(function(k,e,i,j){k.fn.caret=function(b,l){var a,c,f=this[0],d=(navigator.userAgent.toUpperCase().indexOf('MSIE') >= 0);if(typeof b==="object"&&typeof b.start==="number"&&typeof b.end==="number"){a=b.start;c=b.end}else if(typeof b==="number"&&typeof l==="number"){a=b;c=l}else if(typeof b==="string")if((a=f.value.indexOf(b))>-1)c=a+b[e];else a=null;else if(Object.prototype.toString.call(b)==="[object RegExp]"){b=b.exec(f.value);if(b!=null){a=b.index;c=a+b[0][e]}}if(typeof a!="undefined"){if(d){d=this[0].createTextRange();d.collapse(true);
d.moveStart("character",a);d.moveEnd("character",c-a);d.select()}else{this[0].selectionStart=a;this[0].selectionEnd=c}this[0].focus();return this}else{if(d){c=document.selection;if(this[0].tagName.toLowerCase()!="textarea"){d=this.val();a=c[i]()[j]();a.moveEnd("character",d[e]);var g=a.text==""?d[e]:d.lastIndexOf(a.text);a=c[i]()[j]();a.moveStart("character",-d[e]);var h=a.text[e]}else{a=c[i]();c=a[j]();c.moveToElementText(this[0]);c.setEndPoint("EndToEnd",a);g=c.text[e]-a.text[e];h=g+a.text[e]}}else{g=
f.selectionStart;h=f.selectionEnd}a=f.value.substring(g,h);return{start:g,end:h,text:a,replace:function(m){return f.value.substring(0,g)+m+f.value.substring(h,f.value[e])}}}}})(jQuery,"length","createRange","duplicate");

function ($) {

	$.fn.extend({
		youqi_datetime: function (options) {

			var defaults = {
				dateOnly: false,
				padding: 20,
				mouseOverColor: '#000000',
				mouseOutColor: '#ffffff'
			}

			var options = $.extend(defaults, options);

			//Iterate over the current set of matched elements
			return this.each(function () {
				var o = options;
				$(this).val(getUNIXTimeNow());

				$(this).focus(function (event) {
					event.stopPropagation;
					return false;

				}).mousedown(function () {

				}).blur(function () {
					var max_loop = 50;
					if ($(this).val() == "") {
						$(this).val(getUNIXTimeNow());
					}

					while (!$(this).val().match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/) && --max_loop > 0) {
						var temp_value = $(this).val().replace(/[^0-9]/g, "").replace(/$/, "00000000000000");

						if (temp_value.match(/^\d+$/)) {
							$(this).val(temp_value.substr(0, 4) + "-" + temp_value.substr(4, 2) + "-" + temp_value.substr(6, 2) + " " + temp_value.substr(8, 2) + ":" + temp_value.substr(10, 2) + ":" + temp_value.substr(12, 2));
							continue;
						}

						if ($(this).val().match(/[^0-9: -]+/)) {
							$(this).val($(this).val().replace(/[^0-9: -]+/, ""));
							continue;
						}

						if ($(this).val().match(/^\d{4}-\d{2}-\d{2}\s*$/)) {
							$(this).val($(this).val().replace(/^(\d{4}-\d{2}-\d{2}).*/, "$1 00:00:00"));
							continue;
						}

						if ($(this).val().match(/^\d{4}-\d{2}-\d{2}\s*\d{2}:\d{2}:?\s*$/)) {
							$(this).val($(this).val().replace(/^(\d{4}-\d{2}-\d{2})\s*(\d{2}:\d{2}).*/, "$1 $2:00"));
							continue;
						}

						if ($(this).val().match(/^\d{4}-\d{2}-\d{2}\s*\d{2}:?\s*$/)) {
							$(this).val($(this).val().replace(/^(\d{4}-\d{2}-\d{2})\s*(\d{2}).*/, "$1 $2:00:00"));
							continue;
						}

						if ($(this).val().match(/^\d{4}-\d{2}-\d{2}\s*\d{4}:?\s*$/)) {
							$(this).val($(this).val().replace(/^(\d{4}-\d{2}-\d{2})\s*(\d{2})(\d{2}).*/, "$1 $2:$3:00"));
							continue;
						}

						if ($(this).val().match(/^\d{4}-\d{2}-\d{2}\s*\d{6}\s*$/)) {
							$(this).val($(this).val().replace(/^(\d{4}-\d{2}-\d{2})\s*(\d{2})(\d{2})(\d{2}).*/, "$1 $2:$3:$4"));
							continue;
						}

						if ($(this).val().match(/^\d{4}-\d{2}-\d{2}\s*\d{6}\s*$/)) {
							$(this).val($(this).val().replace(/^(\d{4}-\d{2}-\d{2})\s*(\d{2})(\d{2})(\d{2}).*/, "$1 $2:$3:$4"));
							continue;
						}

						//remove spaces
						if ($(this).val().match(/\s+$/)) {
							$(this).val($(this).val().replace(/\s+$/, ""));
							continue;
						}

						//UNFIXABLE...
						$(this).val(getUNIXTimeNow());
					}
				}).keydown(function (event) {
					//event.stopPropagation();
					if ($(this).val().length == 19 && $(this).val().match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
						//return false;
					}

					if (event.which == 38) //up
					{
						if (!$(this).val().match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
							return true;
						}
						var year = $(this).val().substring(0, 4);
						var month = $(this).val().substring(5, 7);
						var day = $(this).val().substring(8, 10);

						var d = new Date(year, (month - 1), day);
						d.setTime(d.getTime() - 86400000);
						var yesterday = d.getFullYear() + "-" + ((d.getMonth() + 1) < 10 ? ("0" + (d.getMonth() + 1)) : (d.getMonth() + 1)) + "-" + (d.getDate() < 10 ? ("0" + d.getDate()) : d.getDate());

						$(this).val(yesterday + $(this).val().substring(10));
						//$(this).setSelectionRange(1, 2);

						focusTime($(this));
						return false;
					}

					if (event.which == 40) //down
					{
						if (!$(this).val().match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
							return true;
						}

						var year = $(this).val().substring(0, 4);
						var month = $(this).val().substring(5, 7);
						var day = $(this).val().substring(8, 10);

						var d = new Date(year, (month - 1), day);
						d.setTime(d.getTime() + 86400000);
						var tomorrow = d.getFullYear() + "-" + ((d.getMonth() + 1) < 10 ? ("0" + (d.getMonth() + 1)) : (d.getMonth() + 1)) + "-" + (d.getDate() < 10 ? ("0" + d.getDate()) : d.getDate());

						$(this).val(tomorrow + $(this).val().substring(10));

						focusTime($(this));
						return false;
					}
					//alert('Handler for .keypress() called.' + event.which );
				});


			});
		}
	});

	function getUNIXTimeNow() {
		var d = new Date();
		return d.getUTCFullYear() + "-" +
			zeroPad(d.getUTCMonth() + 1, 2) + "-" +
			zeroPad(d.getUTCDate(), 2) + " " +
			zeroPad(d.getHours(), 2) + ":" +
			zeroPad(d.getMinutes(), 2) + ":" +
			zeroPad(d.getSeconds(), 2)
	}

	function focusTime(objInput) {
		objInput.caret({
			start: 11,
			end: 19
		});
	}

	function zeroPad(num, count) {
		var numZeropad = num + '';
		while (numZeropad.length < count) {
			numZeropad = "0" + numZeropad;
		}
		return numZeropad;
	}
})(jQuery);