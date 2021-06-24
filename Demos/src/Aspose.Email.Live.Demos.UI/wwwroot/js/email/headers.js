(function ($) {
    $.headers = function (data) {

        var mainDiv = $('<div class="col-md-12" />');

        mainDiv.append($('<br/>'));

        if (data.Traces === null || data.Traces.length === 0) {
            mainDiv.append($('<br/><br/><p>There is no data in this email</p><br/>'));
        }
        else {
            mainDiv.append($('<p>Email Servers:</p>'));
        }

        var div = $('<div class="list-group rawHeaders"/>');
        mainDiv.append(div);

        for (var i in data.Traces) {

            var item = data.Traces[i];
            var domain = item.Domain;
            var ip = item.Ip;
            var location = item.Location;
          
            var itemDiv = $('<div class="list-group-item text-left"/>');

            if (domain === null)
                domain = "no domain";

            if (ip === null)
                ip = "no ip";

            var h4 = $('<h4 class="list-group-item-heading bg-primary">&nbsp;<b>' + domain + '</b> (' + ip + ')</h4>');
            itemDiv.append(h4);

                var ipInfo = $('<div class="list-group-item-text" id="ipInfo"/>');
                itemDiv.append(ipInfo);

                var ipTable = $('<table class="table"/>');
                ipInfo.append(ipTable);

            if (location !== null &&
                location.CountryCode !== null && location.CountryCode !== '' &&
                location.Country !== null && location.Country !== '' &&
                location.Region !== null && location.Region !== '' &&
                location.City !== null && location.City !== '') {

                var tHead = $('<thead><tr><th>CountryCode</th><th>Country</th><th>Region</th><th>City</th></tr></thead>');
                ipTable.append(tHead);
                var tbody = $('<tbody/>');
                ipTable.append(tbody);

                var tr = $('<tr/>');
                tbody.append(tr);

                tr.append($('<td>' + location.CountryCode + '</td>'));
                tr.append($('<td>' + location.Country + '</td>'));
                tr.append($('<td>' + location.Region + '</td>'));
                tr.append($('<td>' + location.City + '</td>'));
            }

            var whoIs = item.WhoIs;

            if (whoIs !== null) {
                var info = whoIs.Info;
                var error = whoIs.Error;

                var whoisContainer = $('<div style="padding:8px;" id="whoisContainer"/>');
                itemDiv.append(whoisContainer);

                whoisContainer.append($('<b>Whois:</b>'));
                if (info === null) {
                    whoisContainer.append($('<p>' + error + '</p>'));
                }
                else {
                    whoisContainer.append($('<textarea readonly style="width:100%;">' + info + '</textarea>'));
                    itemDiv.append($('<br />'));
                    itemDiv.append($('<br />'));
                }
            }

            div.append(itemDiv);
        }

        mainDiv.append($('<div class="convertbtn"><input type="button" class="btn btn-reset btn-lg" ID="resetButton" value="RESET" /></div>'));

        $('.filedrop').addClass('hidden');
        $('input#uploadButton').addClass('hidden');
        $('div#rawHeadersBlock').addClass('hidden');

        $('div#headers > div').remove();
        $('div#headers').append(mainDiv);

        var reset = function () {
            $('div#headers > div').remove();

            $('.filedrop').removeClass('hidden');
            $('input#uploadButton').removeClass('hidden');
            $('div#rawHeadersBlock').removeClass('hidden');
            fileDrop.reset();
            window.scrollTo(0, 0);
        };

        mainDiv.on("click", ".btn-reset", reset);
    };
})(jQuery);