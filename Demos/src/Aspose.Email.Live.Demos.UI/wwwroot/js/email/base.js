const VIEWABLE_EXTENSIONS = [
    'MSG', 'EML', 'MBOX', 'OST', 'PST'
];


// filedrop component
var fileDrop = {};
var fileDrop2 = {};

var lastErrorInfo = {};

$.extend($.expr[':'], {
    isEmpty: function (e) {
        return e.value === '';
    }
});

// Restricts input for the set of matched elements to the given inputFilter function.
(function ($) {
    $.fn.inputFilter = function (inputFilter) {
        return this.on("input keydown keyup mousedown mouseup select contextmenu drop", function () {
            if (inputFilter(this.value)) {
                this.oldValue = this.value;
                this.oldSelectionStart = this.selectionStart;
                this.oldSelectionEnd = this.selectionEnd;
            } else if (this.hasOwnProperty("oldValue")) {
                this.value = this.oldValue;
                this.setSelectionRange(this.oldSelectionStart, this.oldSelectionEnd);
            } else {
                this.value = "";
            }
        });
    };
}(jQuery));

function showLoader() {
    $('.progress > .progress-bar').html('0%');
    $('.progress > .progress-bar').css('width', '0%');
    $('#loader').removeClass("hidden");
    hideAlert();
}

function hideLoader() {
    $('#loader').addClass("hidden");
}

function generateViewerLink(data) {
    var id = data.FolderName !== undefined ? data.FolderName : data.id;
    return encodeURI(o.ViewerPath +
        'FileName=' +
        data.FileName +
        '&FolderName=' +
        id +
        '&CallbackURL=' +
        o.AppURL);
}

function generateEditorLink(data) {

    var id = data.FolderName !== undefined ? data.FolderName : data.id;

    if (id === undefined)
        return encodeURI(o.EditorPath + 'CallbackURL=' + o.AppURL);

    return encodeURI(o.EditorPath +
        'FileName=' +
        data.FileName +
        '&FolderName=' +
        id +
        '&CallbackURL=' +
        o.AppURL);
}

function sendPageView(url) {
    if ('ga' in window)
        try {
            var tracker = ga.getAll()[0];
            if (tracker !== undefined) {
                tracker.send('pageview', url);
            }
        } catch (e) {
            /**/
        }
}

function workSuccess(data, textStatus, xhr) {
    hideLoader();

    if (data.StatusCode === 200 && data.FileName !== null) {
        if (data.FileProcessingErrorCode !== undefined && data.FileProcessingErrorCode !== 0) {
            lastErrorInfo.folderId = data.FolderName;
            lastErrorInfo.error = o.FileProcessingErrorCodes[data.FileProcessingErrorCode];

            showAlert(o.FileProcessingErrorCodes[data.FileProcessingErrorCode]);
            showReportWindow(o.FileProcessingErrorCodes[data.FileProcessingErrorCode]);
            return;
        }

        $('#WorkPlaceHolder').addClass('hidden');
        $('.appfeaturesectionlist').addClass('hidden');
        $('.sub-menu-container').addClass('hidden');
        $('.overview').addClass('hidden');

        $('.howtolist').addClass('hidden');
        $('.app-features-section').addClass('hidden');
        $('.app-product-section').addClass('hidden');

        $('#DownloadPlaceHolder').removeClass('hidden');
        $('#OtherApps').removeClass('hidden');

        if (o.ReturnFromViewer === undefined) {
            const pos = o.AppDownloadURL.indexOf('?');
            const url = pos === -1 ? o.AppDownloadURL : o.AppDownloadURL.substring(0, pos);
            sendPageView(url);
        }

        var url = encodeURI(o.APIBasePath + `Common/DownloadFile/${data.FolderName}?file=${data.FileName}`);
        $('#DownloadButton').attr('href', url);
        o.DownloadUrl = url;

        if (o.ShowViewerButton) {
            let viewerlink = $('#ViewerLink');
            let dotPos = data.FileName.lastIndexOf('.');
            let ext = dotPos >= 0 ? data.FileName.substring(dotPos + 1).toUpperCase() : null;
            if (ext !== null && viewerlink.length && VIEWABLE_EXTENSIONS.indexOf(ext) !== -1) {
                viewerlink.on('click', function (evt) {
                    evt.preventDefault();
                    evt.stopPropagation();
                    openIframe(generateViewerLink(data), '/email/viewer', '/email/view');
                });
            }
            else {
                viewerlink.hide();
                $(viewerlink[0].parentNode.previousElementSibling).hide(); // div.clearfix
            }
        }
    } else {
        lastErrorInfo.folderId = data.FolderName;
        lastErrorInfo.error = data.Status;

        showAlert(data.Status);
        showReportWindow(data.Status);
    }
}

function sendEmail() {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    var email = $('#EmailToInput').val();
    if (!email || !re.test(String(email).toLowerCase())) {
        window.alert('Please specify the valid email address!');
        return;
    }

    var data = {
        appname: o.AppName,
        email: email,
        url: o.DownloadUrl,
        title: $('#ProductTitle')[0].innerText
    };

    $('#sendEmailModal').modal('hide');
    $('#sendEmailButton').addClass('hidden');

    let url = o.APIBasePath + 'Common/sendemail';

    $.ajax({
        method: 'POST',
        url: url,
        data: JSON.stringify(data),
        contentType: "application/json",
        dataType: 'json',
        success: (data) => {
            showMessage(data.message);
        },
        complete: () => {
            $('#sendEmailButton').removeClass('hidden');
            hideLoader();
        },
        error: handleAjaxError
    });
}

function sendFeedback(text) {
    var msg = (typeof text === 'string' ? text : $('#feedbackText').val());
    if (!msg || msg.match(/^\s+$/) || msg.length > 1000) {
        return;
    }

    if (!text) {
        if ('ga' in window) {
            try {
                var tracker = window.ga.getAll()[0];
                if (tracker !== undefined) {
                    tracker.send('event', {
                        'eventCategory': 'Social',
                        'eventAction': 'feedback-in-download'
                    });
                }
            } catch (e) { }
        }
    }

    let url = o.APIBasePath + 'Common/sendfeedback';

    $.ajax({
        method: "POST",
        url: url,
        data: JSON.stringify({
            'appname': o.AppName,
            'text': msg
        }),
        contentType: "application/json",
        dataType: 'json',
        success: (data) => {
            showMessage(data.message);
            $('#feedback').hide();
        },
        error: handleAjaxError
    });
}

function hideAlert() {
    $('#alertMessage').addClass("hidden");
    $('#alertMessage').text("");
    $('#alertSuccess').addClass("hidden");
    $('#alertSuccess').text("");
}

function showAlert(msg) {
    hideLoader();
    $('#alertMessage').html(msg);
    $('#alertMessage').removeClass("hidden");
    $('#alertMessage').fadeOut(100).fadeIn(100).fadeOut(100).fadeIn(100);
}

function showMessage(msg) {
    hideLoader();
    $('#alertSuccess').text(msg);
    $('#alertSuccess').removeClass("hidden");
}

(function ($) {
    $.QueryString = (function (paramsArray) {
        let params = {};

        for (let i = 0; i < paramsArray.length; ++i) {
            let param = paramsArray[i]
                .split('=', 2);

            if (param.length !== 2)
                continue;

            params[param[0]] = decodeURIComponent(param[1].replace(/\+/g, " "));
        }

        return params;
    })(window.location.search.substr(1).split('&'))
})(jQuery);

function progress(evt) {
    if (evt.lengthComputable) {
        var max = evt.total;
        var current = evt.loaded;

        var percentage = Math.round((current * 100) / max);
        percentage = percentage + '%';

        $('.progress > .progress-bar').html(percentage);
        $('.progress > .progress-bar').css('width', percentage);
    }
}

function removeAllFileBlocks() {
    fileDrop.droppedFiles.forEach(function (item) {
        $('#fileupload-' + item.id).remove();
    });
    fileDrop.droppedFiles = [];
    hideLoader();
}

function openIframe(url, fakeUrl, pageViewUrl) {
    // push fake state to prevent from going back
    window.history.pushState(null, null, fakeUrl);

    // remove body scrollbar
    $('body').css('overflow-y', 'hidden');

    // create iframe and add it into body
    var div = $('<div id="iframe-wrap"></div>');
    $('<iframe>', {
        src: url,
        id: 'iframe-document',
        frameborder: 0,
        scrolling: 'yes'
    }).appendTo(div);
    div.appendTo('body');
    sendPageView(pageViewUrl);
}

function closeIframe() {
    removeAllFileBlocks();
    $('div#iframe-wrap').remove();
    $('body').css('overflow-y', 'auto');
}


function request(url, data) {
    showLoader();
    $.ajax({
        method: 'POST',
        url: url,
        data: data,
        processData: false,
        contentType: false,
        cache: false,
        timeout: 1000000,
        success: workSuccess,
        xhr: function () {
            var myXhr = $.ajaxSettings.xhr();
            if (myXhr.upload)
                myXhr.upload.addEventListener('progress', progress, false);
            return myXhr;
        },
        error: handleAjaxError
    });
}

function requestMerger() {
    let data = fileDrop.prepareFormData(2, o.MaximumUploadFiles);
    if (data === null)
        return;

    let url = o.APIBasePath + 'AsposeEmailMerger/Merge?outputType=' + $('#saveAs').val();
    request(url, data);
}

function requestParser() {
    let data = fileDrop.prepareFormData();
    if (data === null)
        return;

    let url = o.APIBasePath + 'AsposeEmailParser/Parse';
    request(url, data);
}

function requestDisposable() {
    var options = o;
    var email = $('#DisposableQuery').val();
    if (email == null || email.length < 1) {
        showAlert(options.EmailInputMessage);
        return;
    }
    var data = new FormData();
    data.append('email', email);


    //let data = {
    //	email: email
    //};

    let url = o.APIBasePath + 'AsposeEmailDisposable/IsDisposable';
    requestPostDisposable(url, data);
}

function requestPostDisposable(url, data) {
    var options = o;
    showLoader();
    $.ajax({
        method: 'POST',
        url: url,
        //dataType: 'json',
        //data: JSON.stringify(data),
        data: data,
        processData: false,
        contentType: false,
        cache: false,
        timeout: 1000000,
        success: function (workData, textStatus, xhr) {
            if (workData)
                showMessage(options.EmailDisposableMessage);
            else
                showMessage(options.EmailNotDisposableMessage)
        },
        xhr: function () {
            var myXhr = $.ajaxSettings.xhr();
            if (myXhr.upload)
                myXhr.upload.addEventListener('progress', progress, false);
            return myXhr;
        },
        error: handleAjaxError
    });
}

function requestPasswordStrength() {
    var options = o;
    var password = $('#ugpassword').val();
    if (password == null || password.length < 1) {
        showAlert(options.PasswordInputMessage);
        return;
    }
    var data = new FormData();
    data.append('password', password);

    let url = o.APIBasePath + 'AsposeEmailPasswordStrength/CheckStrength';
    requestPostPasswordStrength(url, data);
}

function requestPostPasswordStrength(url, data) {
    var options = o;
    showLoader();
    $.ajax({
        method: 'POST',
        url: url,
        //dataType: 'json',
        //data: JSON.stringify(data),
        data: data,
        processData: false,
        contentType: false,
        cache: false,
        timeout: 1000000,
        success: function (workData, textStatus, xhr) {
            $(function () {
                var div = $("<div/>");
                var result = $("<result/>");
                var table = $('<table class="admin_mGrid" cellspacing="0" border="0" style="border-collapse:collapse; margin-bottom:50px;" />'),
                    table_head = $('<thead/>'),
                    head_row = $('<tr/>'),
                    table_body = $('<tbody/>');

                head_row
                    .append($('<th>').text(""))
                    .append($('<th>').text(options.Keys["PasswordStrengthKey"]))
                    .append($('<th>').text(options.Keys["PasswordStrengthCount"]))
                    .append($('<th>').text(options.Keys["PasswordStrengthBonus"]));

                var score = 0;
                $.each(workData, function (i, item) {
                    score += item.Bonus;
                    var $tr =
                        $('<tr>')
                            .append($('<td>').text(i + 1))
                            .append($('<td>').text(options.Keys[item.NameKey]))
                            .append($('<td>').text(item.Count))
                            .append($('<td>').text(item.Bonus))
                        ;

                    var color = "";
                    switch (item.Rate) {
                        case 1:
                            color = "blue";
                            break;

                        case 2:
                            color = "green";
                            break;

                        case 3:
                            color = "sandybrown";
                            break;

                        case 4:
                            color = "red";
                            break;
                    }
                    
                    $tr.attr('style', "font-weight: bold;color:" + color);
                    table_body.append($tr);
                });

                table_head.append(head_row);
                table.append(table_head);
                table.append(table_body);

                result
                    .append($('<label>').text(options.Keys["PasswordScore"] + ' ' + score));

                div
                    .append(result)
                    .append(table);

                $('#resultCheck').html(div);
            });

            hideAlert();
            hideLoader();
        },
        xhr: function () {
            var myXhr = $.ajaxSettings.xhr();
            if (myXhr.upload)
                myXhr.upload.addEventListener('progress', progress, false);
            return myXhr;
        },
        error: handleAjaxError
    });
}


function requestHeaders() {
    let data = fileDrop.prepareFormData();

    var rawHeaders = $('#txtHeadersAsText').val();
    var options = o;
    if (data === null && !rawHeaders.length) {
        showAlert(options.FileSelectMessage);
        return;
    }
    hideAlert();
    if (data === null) {
        data = new FormData();

        rawHeaders = rawHeaders.replace(/^\s*[\r\n]/gm, "");
        console.log(rawHeaders);

        data.append('txtHeaders', rawHeaders);
    }

    let url = o.APIBasePath + 'AsposeEmailHeaders/AnalyzeEmailFiles?skipLocal=true';
    showLoader();
    $.ajax({
        method: 'POST',
        url: url,
        data: data,
        processData: false,
        contentType: false,
        cache: false,
        timeout: 1200000,
        success: (d) => {
            hideLoader();
            $.headers(d);
        },
        xhr: function () {
            var myXhr = $.ajaxSettings.xhr();
            if (myXhr.upload)
                myXhr.upload.addEventListener('progress', progress, false);
            return myXhr;
        },
        error: handleAjaxError
    });
}

function requestAnnotation() {
    let data = fileDrop.prepareFormData();
    if (data === null)
        return;

    let url = o.APIBasePath + 'AsposeEmailAnnotation/Remove';
    request(url, data);
}

function requestConversion() {
    let data = fileDrop.prepareFormData();
    if (data === null)
        return;

    let url = o.APIBasePath + 'AsposeEmailConversion/Convert?outputType=' + $('#saveAs').val();
    request(url, data);
}

function requestMetadata(data) {
    let url = o.APIBasePath + 'AsposeEmailMetadata/properties';

    $.ajax({
        method: 'POST',
        url: url,
        data: JSON.stringify(data),
        processData: false,
        contentType: false,
        cache: false,
        timeout: 1200000,
        success: (d) => {
            $.metadata(d, data.id, data.FileName);
        },
        error: handleAjaxError
    });
}

function handleAjaxError(xhr, err, thrown) {
    if (xhr.data !== undefined && xhr.data.Status !== undefined) {
        lastErrorInfo.folderId = xhr.data.FolderName;
        lastErrorInfo.error = xhr.data.Status;

        showAlert(xhr.data.Status);
        showReportWindow(xhr.data.Status);
    }
    else {

        let error = "Error " + xhr.status + ": " + xhr.statusText;

        if (xhr.readyState == 4) {
            error = "Error with HTTP protocol. The file is too big for this operation.";
            showAlert(error);
            return;
        }
        else if (xhr.readyState == 0) {
            if (xhr.statusText === "timeout") {
                error = "Request timeout, connection refused.";
                showAlert(error);
                return;
            }
            if (xhr.statusText === "error") {
                error = "Network error, connection refused. Please check your internet.";
                showAlert(error);
                return;
            }
        }

        lastErrorInfo.folderId = xhr.FolderName;
        lastErrorInfo.error = error;

        showAlert(error);
        showReportWindow(error);
    }
}

function requestRedaction() {
    if (!validateSearch())
        return;

    let data = fileDrop.prepareFormData();
    if (data === null)
        return;

    let url = o.APIBasePath + 'AsposeEmailRedaction/Redact?outputType=' + $('#saveAs').val() +
        '&searchQuery=' + encodeURI($('#searchQuery').val()) +
        '&replaceText=' + encodeURI($('#replaceText').val()) +
        '&caseSensitive=' + $('#caseSensitive').prop('checked') +
        '&text=' + $('#text').prop('checked') +
        '&comments=' + $('#comments').prop('checked') +
        '&metadata=' + $('#metadata').prop('checked');
    request(url, data);
}

function validateSearch() {
    if ($("#searchQuery").val().length)
        return true;
    showAlert(o.validationSearchMessage);
    return false;
}

function requestSearch() {
    if (!validateSearch())
        return;

    let data = fileDrop.prepareFormData();
    if (data === null)
        return;

    let url = o.APIBasePath + 'AsposeEmailSearch/Search?query=' + encodeURI($('#searchQuery').val());
    request(url, data);
}

function validateUnlock() {
    if ($("#passw").val().length)
        return true;
    showAlert(o.validationMessage);
    return false;
}

function validateComparison() {
    if (fileDrop.droppedFiles.length === 1 && fileDrop.droppedFiles.length === 1)
        return true;
    showAlert(o.FileSelectMessage);
    return false;
}

function requestComparison() {
    if (!validateComparison())
        return;

    let data = fileDrop.prepareFormData();
    if (data === null)
        return;

    let data2 = fileDrop2.prepareFormData();
    if (data2 === null)
        return;

    for (var entry of data2.entries())
        data.append(entry[0], entry[1]);

    let url = o.APIBasePath + 'AsposeEmailComparison/Compare';
    request(url, data);
}


function requestViewer(data) {
    fileDrop.reset();
    var url = generateViewerLink(data);
    openIframe(url, o.AppRoute, '/email/view');
}

function requestEditor(data) {
    fileDrop.reset();
    var url = generateEditorLink(data);
    openIframe(url, o.AppRoute, '/email/edit');
}

function prepareDownloadUrl() {
    o.AppDownloadURL = o.AppURL;
    var pos = o.AppDownloadURL.indexOf(':');
    if (pos > 0)
        o.AppDownloadURL = (pos > 0 ? o.AppDownloadURL.substring(pos + 3) : o.AppURL) + '/download';
    pos = o.AppDownloadURL.indexOf('/');
    o.AppDownloadURL = o.AppDownloadURL.substring(pos);
}

function checkReturnFromViewer() {
    var query = window.location.search;
    if (query.length > 0) {
        o.ReturnFromViewer = true;
        var data = {
            StatusCode: 200,
            FolderName: $.QueryString['id'],
            FileName: $.QueryString['FileName'],
            FileProcessingErrorCode: 0
        };
        var beforeQueryString = window.location.href.split("?")[0];
        window.history.pushState({}, document.title, beforeQueryString);

        if (!o.UploadAndRedirect)
            workSuccess(data);
    }
}

function sendFeedbackExtended() {
    var text = $('#feedbackBody').val();
    if (text && !text.match(/^.s+$/)) {
        $('#feedbackModal').modal('hide');
        sendFeedback(text);
    }
}

function otherAppClick(name, left = false) {
    if ('ga' in window) {
        try {
            var tracker = window.ga.getAll()[0];
            if (tracker !== undefined) {
                tracker.send('event', {
                    'eventCategory': 'Other App Click' + (left ? ' Left' : ''),
                    'eventAction': name
                });
            }
        } catch (e) { }
    }
}

function showReportWindow(error) {

    $('#reportErrorText').text(error);

    $('#reportWindow').removeClass("hidden");
    $('#reportWindow').animate({
        opacity: 1,
        transition: "opacity 0.3s ease-out",
        top: "50%"
    }, 300);

    showDarkMask();
}

function showReportSentWindow(link, message) {

    closeReportWindow();

    $('#linkToForums').attr("href", link);
    $('#reportSuccessLabel').text(message);

    $('#reportSentWindow').removeClass("hidden");
    $('#reportSentWindow').animate({
        opacity: 1,
        transition: "opacity 0.3s ease-out",
        top: "50%"
    }, 300);
}

function closeReportWindow() {
    $('#reportWindow').animate({
        opacity: 0,
        transition: "opacity 0.3s ease-out",
        top: "60%"
    }, 300, 'linear', function () {
        $('#reportWindow').addClass('hidden');
    });
}

function closeReportSentWindow() {
    $('#reportSentWindow').animate({
        opacity: 0,
        transition: "opacity 0.3s ease-out",
        top: "60%"
    }, 300, 'linear', function () {
        $('#reportSentWindow').addClass('hidden');
    });
}

function showDarkMask() {
    $('#reportWindowMask').removeClass("hidden");
    $('#reportWindowMask').animate({
        opacity: .8,
        transition: "opacity 0.3s ease-out"
    }, 300);
}

function removeDarkMask() {
    $('#reportWindowMask').animate({
        opacity: 0,
        transition: "opacity 0.3s ease-out"
    }, 300, 'linear', function () {
        $('#reportWindowMask').addClass('hidden');
    });
}

function trySendReport(e) {
    let email = $('#feedbackEmail').val();
    let isPrivate = $('#privateReportCheckBox').val() === "on";

    let data = {
        'folderId': lastErrorInfo.folderId,
        'errorText': lastErrorInfo.error,
        'appname': o.AppName,
        'email': email,
        'isPrivate': isPrivate
    };

    let url = o.APIBasePath + 'Common/sendErrorReport';

    $.ajax({
        method: 'POST',
        url: url,
        data: JSON.stringify(data),
        contentType: 'application/json',
        cache: false,
        timeout: 30000,
        success: (d) => {
            showMessage('The report was successfully sent');
            showReportSentWindow(d.link, d.message);
        },
        error: (err) => {
            if (err.data !== undefined && err.data.Status !== undefined)
                showAlert(err.data.Status);
            else
                showAlert("Error " + err.status + ": " + err.statusText);
        }
    });
}

$(document).ready(function () {
    prepareDownloadUrl();
    //checkReturnFromViewer();

    fileDrop = $('form#UploadFile').filedrop(Object.assign({
        showAlert: showAlert,
        hideAlert: hideAlert,
        showLoader: showLoader,
        progress: progress
    }, o));

    if (o.AppName === "Comparison") {
        fileDrop2 = $('form#UploadFile').filedrop(Object.assign({
            showAlert: showAlert,
            hideAlert: hideAlert,
            showLoader: showLoader,
            progress: progress
        }, o));
    }

    //showDarkMask();
    //showReportSentWindow('https://google.com/', 'test message');

    $('#sendReportButton').click(trySendReport);

    $('#closeReportSuccessButton').click(function (e) {
        closeReportSentWindow();
        removeDarkMask();
    });


    $('#closeReportButton').click(function (e) {
        closeReportWindow();
        removeDarkMask();
    })

    // close iframe if it was opened
    window.onpopstate = function (event) {
        if ($('div#iframe-wrap').length > 0) {
            closeIframe();
        }
    };

    if (!o.UploadAndRedirect || o.NeedsProcessButton) {
        $('#uploadButton').on('click', o.Method);
    }

    // social network modal
    $('#bookmarkModal').on('show.bs.modal', function (e) {
        $('#bookmarkModal').css('display', 'flex');
        $('#bookmarkModal').on('keydown', function (evt) {
            if ((evt.metaKey || evt.ctrlKey) && String.fromCharCode(evt.which).toLowerCase() === 'd') {
                $('#bookmarkModal').modal('hide');
            }
        });
    });
    $('#bookmarkModal').on('hidden.bs.modal', function (e) {
        $('#bookmarkModal').off('keydown');
    });

    // send email modal
    $('#sendEmailButton').on('click', function () {
        $('#sendEmailModal').modal({
            keyboard: true
        });
    });
    $('#sendEmailModal').on('show.bs.modal', function () {
        $('#sendEmailModal').css('display', 'flex');
    });
    $('#sendEmailModal').on('shown.bs.modal', function () {
        $('#EmailToInput').focus();
    });

    // send feedback modal
    $('#feedbackModal').on('show.bs.modal', function (e) {
        $('#feedbackModal').css('display', 'flex');
    });
    $('#feedbackModal').on('shown.bs.modal', function () {
        $('#feedbackBody').focus();
    });

    $('#sendFeedbackBtn').on('click', sendFeedback);


    var keys = o.InvalidResourceKeys;

    if (keys && keys.length) {
        for (var i = 0; i < keys.length; i++) {
            console.warn("Resource Key not found: " + keys[i]);
        }
    }

});
