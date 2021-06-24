var fileName = undefined;
var folderName = undefined;
var createNewFile = false;
var featureName = '';
var callbackURL = '';
var apiURL = '/AsposeAPI';
var currentPageCount = 1;
var totalPages = 1;
var currentSelectedPage = 1;
var customPageList = [];
var IsLoadingGif = false;
var prevoiusIndx = -1;
var imagedata = [];
var ZoomValue = 1.00;
var ZoomMode = "zoom";
var PageWidth = '65';

function UpdatePager() {
	//if (customPageList.length > 1) {
	//	ZoomPages(ZoomValue);
	//}

	TotalDocumentPages = totalPages;
	document.getElementById('spantoolpager').innerHTML = currentSelectedPage + ' of ' + totalPages;
	document.getElementById('inputcurrentpage').value = '';

	for (var i = 0; i <= totalPages; i++) {
		var element = document.getElementById('imgt-page-' + i);
		if (element !== undefined && element !== null) {
			element.className = 'page-aside';
			if (i === currentSelectedPage) {
				element.className = 'selectedthumbnail page-aside';
				element.scrollIntoView();
			}
		}
	}
	//console.log('in at end of UpdatePager();');
}

function ZoomPages(zoomOption) {
	ZoomValue = parseFloat(ZoomValue).toFixed(2);
	ZoomMode = "zoom";
	var style = null;
	if (zoomOption === '+') {
		if (ZoomValue >= 3.00) {
			ZoomValue = parseFloat(ZoomValue) + (0.50);
		}
		else if (ZoomValue >= 0.25) {
			ZoomValue = parseFloat(ZoomValue) + (0.25);
		}
		else {
			ZoomValue = parseFloat(ZoomValue) + (0.05);
		}
	}
	else if (zoomOption === '-') {
		if (ZoomValue >= 3.00) {
			ZoomValue -= 0.50;
		}
		else if (ZoomValue > 0.25) {
			ZoomValue -= 0.25;
		}
		else {
			ZoomValue -= 0.05;
		}
	}
	else if (zoomOption === 'w') {
		style = "width: 100%;";
		ZoomMode = "width";
	}
	else if (zoomOption === 'h') {
		var pagesContainer = document.getElementsByName("dvPages")[0];
		style = pagesContainer.clientHeight;
		ZoomMode = "height";
	}
	else {
		ZoomValue = parseFloat(zoomOption);
	}

	if (ZoomValue < 0.05) {
		ZoomValue = 0.05;
	}
	else if (ZoomValue > 6) {
		ZoomValue = 6;
	}

	//console.log('Before Parse ZoomValue: ' + ZoomValue);
	ZoomValue = parseFloat(ZoomValue).toFixed(2);
	//console.log('After ZoomValue: ' + ZoomValue);
	var lstdvPages = document.getElementsByName("dvInnerPages");

	for (var i = 0; i < lstdvPages.length; i++) {
		//document.getElementById("img-page-" + (i + 1)).style.cssText = "width: 100%;";
		const page = lstdvPages[i];
		const img = page.getElementsByTagName('img')[0];

		var w, h;
		if (style) {
			if (zoomOption === 'h') {
				h = style;
				var ratio = 1;
				if (img.naturalWidth && img.naturalHeight) {
					ratio = img.naturalWidth / img.naturalHeight;
				}
				else if (page.imageWidth && page.imageHeight) {
					ratio = page.imageWidth / page.imageHeight;
				}
				w = Math.round(h * ratio);
				page.style.cssText = `width: ${w}px; height: ${h}px;`;
			} else {
				page.style.cssText = style;
			}
		} else {
			if (img.naturalWidth && img.naturalHeight) {
				w = img.naturalWidth;
				h = img.naturalHeight;
			}
			else if (page.imageWidth && page.imageHeight) {
				w = page.imageWidth;
				h = page.imageHeight;
			}
			else {
				w = 640;
				h = 480;
			}
			w = Math.round(w * ZoomValue);
			h = Math.round(h * ZoomValue);
			page.style.cssText = `width: ${w}px; height: ${h}px;`;
		}

		//if (ZoomValue > 1.40 || ZoomValue < 0.70) {
		//	lstdvPages[i].style.cssText = "float: left; border: 1px solid #058cda; margin: 10px; border-radius: 2px; height: auto; width: " + (100 * ZoomValue) + "%;";
		//}
		//else {
		//	lstdvPages[i].style.cssText = "float: left; border: 1px solid #058cda; margin: 10px; border-radius: 2px;";
		//}
	}
}

function GetPageCss(pageIndex) {
	var lstdvPages = document.getElementsByName("dvInnerPages");
	if (lstdvPages.length <= pageIndex) {
		return "";
	}

	const page = lstdvPages[pageIndex];
	const img = page.getElementsByTagName('img')[0];

	var w, h;

	switch (ZoomMode) {
		case "width":
			return "width: 100%;";
		case "height":
			{
				var pagesContainer = document.getElementsByName("dvPages")[0];
				h = pagesContainer.clientHeight;
				var ratio = 1;
				if (img && img.naturalWidth && img.naturalHeight) {
					ratio = img.naturalWidth / img.naturalHeight;
				}
				else if (page.imageWidth && page.imageHeight) {
					ratio = page.imageWidth / page.imageHeight;
				}
				w = Math.round(h * ratio);
				return  `width: ${w}px; height: ${h}px;`;
			}
		default:
			{
				if (img && img.naturalWidth && img.naturalHeight) {
					w = img.naturalWidth;
					h = img.naturalHeight;
				}
				else if (page.imageWidth && page.imageHeight) {
					w = page.imageWidth;
					h = page.imageHeight;
				}
				else {
					w = 640;
					h = 480;
				}
				w = Math.round(w * ZoomValue);
				h = Math.round(h * ZoomValue);
				return `width: ${w}px; height: ${h}px;`;
			}
	}
}

function SetPageSize(pageIndex, width, height) {
	var lstdvPages = document.getElementsByName("dvInnerPages");
	if (lstdvPages.length > pageIndex) {
		var page = lstdvPages[pageIndex];
		page.imageWidth = width;
		page.imageHeight = height;
	}
}