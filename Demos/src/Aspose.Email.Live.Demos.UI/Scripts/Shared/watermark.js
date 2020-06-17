var pickerText = null;

var linkTabs = function (target) {
    var tabButtons = typeof target === 'string' ? document.querySelector(target) : target;
    var showTab = function (tabsLinkTarget) {

        var tabsPaneTarget, tabsLinkActive, tabsPaneShow;

		var ref = tabsLinkTarget.getAttribute('href');
		tabsPaneTarget = document.querySelector(ref);
        tabsLinkActive = tabsLinkTarget.parentElement.querySelector('.tabs__link_active');
        tabsPaneShow = tabsPaneTarget.parentElement.querySelector('.tabs__pane_show');

        if (tabsLinkTarget === tabsLinkActive) {
            return;
        }

        // remove active classes
        if (tabsLinkActive !== null) {
            tabsLinkActive.classList.remove('tabs__link_active');
        }
        if (tabsPaneShow !== null) {
            tabsPaneShow.classList.remove('tabs__pane_show');
        }

		o.watermarkType = tabsLinkTarget.getAttribute('watermarkType');
        // add active classes
        tabsLinkTarget.classList.add('tabs__link_active');
        tabsPaneTarget.classList.add('tabs__pane_show');
    };

    var switchTabTo = function (tabsLinkIndex) {

        var tabsLinks = tabButtons.querySelectorAll('.tabs__link');

        if (tabsLinks.length > 0) {
            if (tabsLinkIndex > tabsLinks.length) {
                tabsLinkIndex = tabsLinks.length;
            } else if (tabsLinkIndex < 1) {
                tabsLinkIndex = 1;
            }
            showTab(tabsLinks[tabsLinkIndex - 1]);
        }
    };

    tabButtons.addEventListener('click', function (e) {
        var tabsLinkTarget = e.target;

        if (!tabsLinkTarget.classList.contains('tabs__link')) {
            return;
        }

        e.preventDefault();
        showTab(tabsLinkTarget);
    });

    return {
        showTab: function (target) {
            showTab(target);
        },
        switchTabTo: function (index) {
            switchTabTo(index);
        }
    };
};

function colorPicker(color_picker) {
	var colorList = [
		'000000', '993300', '333300', '003300', '003366', '000066', '333399', '333333',
		'660000', 'FF6633', '666633', '336633', '336666', '0066FF', '666699', '666666', 'CC3333', 'FF9933', '99CC33',
		'669966', '66CCCC', '3366FF', '663366', '999999', 'CC66FF', 'FFCC33', 'FFFF66', '99FF66', '99CCCC', '66CCFF',
		'993366', 'CCCCCC', 'FF99CC', 'FFCC99', 'FFFF99', 'CCffCC', 'CCFFff', '99CCFF', 'CC99FF', 'FFFFFF'
	];

	var picker = $(color_picker);

	for (var i = 0; i < colorList.length; i++) {
		picker.append('<li class="color-item" data-hex="' + '#' + colorList[i] + '" style="background-color:' + '#' + colorList[i] + ';"></li>');
	}

	$('body').click(function () {
		picker.fadeOut();
	});

	picker.siblings('.call-picker').click(function (event) {
		event.stopPropagation();
		picker.fadeIn();
		picker.children('li').hover(function () {
			var codeHex = $(this).data('hex');

			picker.siblings('.color-holder').css('background-color', codeHex);
			picker.siblings('input.call-picker').val(codeHex).trigger('change');
		});
	});
	return picker;
};

function validateWatermark() {
	switch (o.watermarkType) {
		default:
			break;
		case "remove":
			break;
		case "text":
			var t = $('#watermarkText').val();
			if (t.length === 0) {
				showAlert(o.validationTextMessage);
				return false;
			}
			break;
		case "image":
			var files = $('#fileUploadImageInput')[0].files;
			if (files === undefined || files.length === 0) {
				showAlert(o.FileSelectMessage);
				return false;
			}
			break;
	}
	return true;
};

function requestWatermark() {
	if (!validateWatermark())
		return;

	let data = fileDrop.prepareFormData();
	if (data === null)
		return;

	let url = o.UIBasePath + 'api/AsposeEmailWatermark/ProcessWatermark?' + 'type=' + o.watermarkType;

	switch (o.watermarkType) {
		case "remove":
			break;
		case "text":
			data.append('text', $('#watermarkText').val());
			url += "&textColor=" + $('#pickcolorText').val().substring(1);
			break;
		case "image":
			var file = $('#fileUploadImageInput')[0].files[0];
			data.append('image', file, file.name);
			break;
	}

	request(url, data);
}

function removefileimage() {
	$('#fileUploadImage').hide();
	$('#fileUploadImageInput').val(null);
	$('#fileUploadImageInput').show();
}

$(document).ready(function () {
	$('#fileUploadImage').hide();

	$('#fileUploadImageInput').change(function () {
		var file = $('#fileUploadImageInput')[0].files[0].name;
		$('#fileUploadImage label').text(file);
		$('#fileUploadImage').show();
	});

    pickerText = new colorPicker('#color-pickerText');

    linkTabs('.app-contenttabs');
	o.watermarkType = "text";
});