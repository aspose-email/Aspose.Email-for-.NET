var stage = 0;
var delimiter = ',';

function removeTemplateData() {
    $('#fileUploadTemplate').hide();
    $('#fileUploadTemplateInput').val(null);
}

function removeDataSourceData() {
    $('#fileUploadDataSource').hide();
    $('#fileUploadDataSourceInput').val(null);

    $("#TableIndex").hide();
    $("#TableDelimiter").hide();
}

function setDelimiter(newDelimiter, buttontext = null) {
    delimiter = newDelimiter;
    if (buttontext === null)
        buttontext = delimiter;
    $('#delimiter')[0].innerText = buttontext;
};

$(document).ready(function () {

    var stageLabel = $("#StageLabel");

    var helpTemplateButton = $("#TemplateHelpButton");
    var templateFiledrop = $("#TemplateFileDrop");

    var helpDataSourceButton = $("#DataSourceHelpButton");
    var dataSourceFiledrop = $("#DataSourceFileDrop");
    var dataSourceForms = $("#DataSourceForms");

    var nextButton = $("#NextButton");
    var backButton = $("#BackButton");
    var assembleAnotherFileButton = $("#AnotherFileButton");

    var assembleButton = $("#AssembleButton");

    var tableIndexDiv = $("#TableIndex");
    var tableDelimiterDiv = $("#TableDelimiter");

    tableIndexDiv.hide();
    tableDelimiterDiv.hide();

    $('#fileUploadTemplate').hide();

    $('#fileUploadTemplateInput').change(function () {
        var file = $('#fileUploadTemplateInput')[0].files[0].name;
        $('#fileUploadTemplate label').text(file);
        $('#fileUploadTemplate').show();
    });

    $('#fileUploadDataSource').hide();

    $('#fileUploadDataSourceInput').change(function () {
        var file = $('#fileUploadDataSourceInput')[0].files[0].name;

        var re = /(?:\.([^.]+))?$/;
        var extension = re.exec(file)[1].toLowerCase();
        switch (extension) {
            case "xml":
            case "json":
                tableIndexDiv.hide();
                tableDelimiterDiv.hide();
                break;
            case "csv":
                tableIndexDiv.hide();
                tableDelimiterDiv.show();
                break;
            default:
                tableIndexDiv.show();
                tableDelimiterDiv.hide();
                break;
        }

        $('#fileUploadDataSource label').text(file);
        $('#fileUploadDataSource').show();
    });

    function toNextStage() {
        ToStage(stage+1);
    }

    function toBackStage() {
        ToStage(stage-1);
    }

    function reset() {
        ToStage(0);
        removeDataSourceData();
        removeTemplateData();

        tableIndexDiv.val(null);
        tableDelimiterDiv.val(null);
    }

    nextButton.click(function (event) {
        event.preventDefault();
        toNextStage();
    });

    backButton.click(function (event) {
        event.preventDefault();
        toBackStage();
    });

    assembleAnotherFileButton.click(function (event) {
        event.preventDefault();
        reset();
    });

    function ToStage(newStage, animated = true) {
        var stageContent = $("#StageContent");

        if (animated) {
            stageContent.fadeOut(300, function () {
                ToggleStageViews(newStage);
                stageContent.fadeIn();
            });
        }
        else {
            ToggleStageViews(newStage);
        }
    }

    function ToggleStageViews(newStage) {
        stage = newStage;
        switch (newStage) {
            case 0:
				{
					stageLabel.text("STAGE 1/3");
					helpTemplateButton.show();
					templateFiledrop.show();
					helpDataSourceButton.hide();
					dataSourceFiledrop.hide();
					dataSourceForms.hide();

					nextButton.show();
					backButton.hide();
					assembleAnotherFileButton.hide();
					assembleButton.hide();

					break;
				}
            case 1:
                {
                    stageLabel.text("STAGE 2/3");
                    helpTemplateButton.hide();
                    templateFiledrop.hide();
                    helpDataSourceButton.show();
                    dataSourceFiledrop.show();
                    dataSourceForms.show();

                    nextButton.show();
                    backButton.show();
                    assembleAnotherFileButton.hide();
                    assembleButton.hide();
                    break;
                }
            case 2:
                {
                    stageLabel.text("STAGE 3/3");

                    helpTemplateButton.hide();
                    templateFiledrop.hide();
                    helpDataSourceButton.hide();
                    dataSourceFiledrop.hide();
                    dataSourceForms.hide();

                    nextButton.hide();
                    backButton.show();
                    assembleAnotherFileButton.show();
                    assembleButton.show();
                    break;
                }
            default: break;
        }
    }


    ToStage(0, false);

    assembleButton.click(function (event) {
        event.preventDefault();

        let templateFiles = $('#fileUploadTemplateInput')[0].files;

        if (templateFiles.length === 0) {
            showAlert(o.FileSelectMessage);
            return;
        }

        let dataSourceFiles = $('#fileUploadDataSourceInput')[0].files;

        if (dataSourceFiles.length === 0) {
            showAlert(o.FileSelectMessage);
            return;
        }

        let data = new FormData();

        let template = templateFiles[0];
        let dataSource = dataSourceFiles[0];

        data.append('template', template, template.name);
        data.append('dataSource', dataSource, dataSource.name);

        var sourceName = $('#DataSourceName').val();
        var tableIndex = $('#datasourceTableIndex').val();

        if (tableIndex === '')
            tableIndex = 0;

		let url = o.UIBasePath + 'api/AsposeEmailAssembly/Assemble?delimiter=' + delimiter + '&sourceName=' + sourceName + '&tableIndex=' + tableIndex;
        request(url, data);
    });
});