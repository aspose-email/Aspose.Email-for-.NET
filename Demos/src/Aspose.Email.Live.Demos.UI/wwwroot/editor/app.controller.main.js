angular.module('AsposeEditorApp').controller('EditorController', function ($scope, $http, $compile) {

	var attachmentsDict = {};
	var attachmentsIndexCounter = 0;

	var createNewFile = false;
	var initialized = false;
	$scope.trum = null; // jquery object
	$scope.t = null; // trumbowyg component
	$scope.loading = false;

	$scope.allowPreview = false;
	$scope.isPreview = false;
	$scope.refresh = true;
	$scope.pageIndex = 1;
	$scope.pageNumber = 1;

	window.history.pushState(null, null, '/email/editor');

	// 'back' button special handler
	/*var onpopstatePrev = window.onpopstate;
	window.onpopstate = function(event) {
		event.preventDefault();
		event.stopPropagation();
		window.onpopstate = onpopstatePrev;
		window.location.href = '/words/editor';
	};*/

	$scope.OpenFileInput = function () {
		$('#file-input').trigger('click');
    }

	$scope.HandleFileInput = function (element) {

		if (element.files) {
			for (var i = 0; i < element.files.length; i++) {

				let file = element.files[i];
				let data = AddItemInAttachments(file.name, GetExtension(file.name), "Document", true);
				let reader = new FileReader();

				reader.onload = function () {
					data.SourceBase64 = btoa(String.fromCharCode.apply(null, new Uint8Array(this.result)));
				}

				reader.readAsArrayBuffer(file);
			}
		}
	};

	function GetExtension(filename) {
		return filename.substring(filename.lastIndexOf('.') + 1, filename.length) || filename
    }

	$scope.RemoveAttachment = function ($event, index, name) {
		let result = confirm('Remove this attachment "' + name + '"?');

		if (result) {

			let attachment = attachmentsDict[index];

			if (attachment.NeedRemove)
				return;

			if (!attachment.WasAdded) {
				attachment.NeedRemove = true;
			}
			else {
				delete attachmentsDict[index];
			}

			$($event.currentTarget).parent().parent().remove();
		}
	}

	$scope.Init = function () {
		if (initialized === false) {
			initialized = true;

			var defaultButtons = [
				['historyUndo', 'historyRedo'],
				['formatting'],
				['fontfamily'],
				['fontsize'],
				['strong', 'em', 'underline'],
				['foreColor', 'backColor'],
				['superscript', 'subscript'],
				['link'],
				['image'],
				['justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'],
				['unorderedList', 'orderedList'],
				['horizontalRule'],
				['fullscreen']
			];

			var buttons = defaultButtons;

			$scope.trum = $('#editor');
			$scope.trum.trumbowyg({
				autogrow: true,
				fixedBtnPane: true,
				minimalLinks: true,
				urlProtocol: true,
				prefix: "editoraspose-",
				svgPath: "/editor/trumbowyg/icons.svg",
				btnsDef: {
					image: {
						dropdown: ['insertImage', 'base64'],
						ico: 'insertImage'
					},
					togglePreview: {
						fn: function () {
							$scope.TogglePreview();
						},
						tag: 'toggle_preview',
						title: 'Toggle editor/preview mode',
						text: 'Preview',
						hasIcon: false,
						isSupported: function () { return $scope.allowPreview; }
					}
				},
				btns: buttons,
				plugins: {
					fontfamily: {
						fontList: [
							{ name: 'Arial', family: 'Arial, Helvetica, sans-serif' },
							{ name: 'Arial Black', family: '\'Arial Black\', Gadget, sans-serif' },
							{ name: 'Comic Sans', family: '\'Comic Sans MS\', Textile, cursive, sans-serif' },
							{ name: 'Courier New', family: '\'Courier New\', Courier, monospace' },
							{ name: 'Georgia', family: 'Georgia, serif' },
							{ name: 'Impact', family: 'Impact, Charcoal, sans-serif' },
							{ name: 'Lucida Console', family: '\'Lucida Console\', Monaco, monospace' },
							{ name: 'Lucida Sans', family: '\'Lucida Sans Uncide\', \'Lucida Grande\', sans-serif' },
							{ name: 'Palatino', family: '\'Palatino Linotype\', \'Book Antiqua\', Palatino, serif' },
							{ name: 'Tahoma', family: 'Tahoma, Geneva, sans-serif' },
							{ name: 'Times New Roman', family: '\'Times New Roman\', Times, serif' },
							{ name: 'Trebuchet', family: '\'Trebuchet MS\', Helvetica, sans-serif' },
							{ name: 'Verdana', family: 'Verdana, Geneva, sans-serif' }
						]
					},
					fontsize: {
						sizeList: [
							'9',
							'14',
							'18',
							'22',
							'30'
						],
						allowCustomSize: true
					}
				}
			});

			$scope.trum.trumbowyg('html', '');
			$scope.t = $scope.trum.data('trumbowyg'); // Component
			$scope.GetEditorHtml();
		}
	};

	$scope.GetEditorHtml = function () {

		if (folderName === '' && fileName === '') {
			createNewFile = true;

			$scope.trum.trumbowyg('html', '>>> Place your content here <<<');

			$('#page-loading').fadeOut(600);
			$('#htmlloader').hide();
			return;
        }


		var paramList = {
			'folderName': folderName,
			'fileName': fileName
		};

		$http({
			method: 'POST',
			url: asposeEditorAPI + 'GetDocumentData',
			params: paramList,
			responseType: "application/json"
		}).then(function (response) {

			let data = response.data;

			$scope.trum.trumbowyg('html', data.Html);

			if (data.Attachments) {

				for (var i = 0; i < data.Attachments.length; i++) {
					let item = data.Attachments[i];
					AddItemInAttachments(item.Name, item.Extension, item.Type, false);
				}
			}

		}, function (error) {
			$scope.ShowError(error.data);
			$scope.trum.trumbowyg('html', '');
		}).finally(function () {
			$('#page-loading').fadeOut(600);
			$('#htmlloader').hide();
		});
	};

	function AddItemInAttachments(name, ext, type, wasAdded) {

		let imageSource = getIconSourceFromExtension(ext);
		let index = attachmentsIndexCounter++;

		let data = {
			Index: index,
			Name: name,
			WasAdded: wasAdded,
			NeedRemove: false,
			Extension: ext
		};

		attachmentsDict[index] = data;

		$('#attachments').append($compile('<div class="col-md-4 media-item">' +
			'<div class="media-content">' +
			'<figure>' +
			'<img class="img-fluid" src="' + imageSource + '"/>' +
			'</figure>' +
			'<div class="media-body col-md-2">' +
			'<h5 class="mt-1 font-weight-bold">' + type + '</h5>' + name +
			'</div>' +
			'<button type="button" class="close" aria-label="Close" ng-click="RemoveAttachment($event, ' + index + ',\'' + name + '\')">' +
			'	<span aria-hidden="true">&times;</span>' +
			'</button>' +
			'</div>' +
			'</div>')($scope));

		return data;
    }
	
	$scope.TogglePreview = function () {

		$scope.isPreview = !$scope.isPreview;
		if ($scope.isPreview) {
			document.activeElement.blur(); // 
			$scope.DoRefresh();
			$('#preview_pane').removeClass('hidden');
			$('#editor_pane').addClass('hidden');
		} else {
			$('#editor_pane').removeClass('hidden');
			$('#preview_pane').addClass('hidden');
		}
	}

	function getIconSourceFromExtension(ext) {
		switch (ext) {

			case ".eml": return "/img/icon-eml.png";
			case ".msg": return "/img/icon-msg.png";
			case ".mbox": return "/img/icon-pst.png";
			case ".ost": return "/img/icon-ost.png";
			case ".pst": return "/img/icon-pst.png";
			case ".pdf": return "/img/icon-pdf.png";
			case ".html": return "/img/icon-html.png";
			case ".pptx": return "/img/icon-pptx.png";
			case ".mht": return "/img/icon-mht.png";
			case ".ppt": return "/img/icon-ppt.png";
			case ".odt": return "/img/icon-doc.png";
			case ".ott": return "/img/icon-doc.png";
			case ".dotx": return "/img/icon-doc.png";
			case ".dotm": return "/img/icon-doc.png";
			case ".docx": return "/img/icon-doc.png";
			case ".docm": return "/img/icon-doc.png";
			case ".doc": return "/img/icon-doc.png";
			case ".rtf": return "/img/icon-doc.png";
			case ".txt": return "/img/icon-text.png";
			case ".xls": return "/img/icon-xls.png";
			case ".jpg": return "/img/icon-jpg.png";
			case ".jpeg": return "/img/icon-jpg.png";
			case ".png": return "/img/icon-image.png";
			case ".tiff": return "/img/icon-tiff.png";
			case ".bmp": return "/img/icon-bmp.png";
			case ".tga": return "/img/icon-image.png";
			case ".svg": return "/img/icon-svg.png";
			case ".ico": return "/img/icon-image.png";
			default:
				return "/img/icon-doc.png";
        }
	}

	$scope.DoRefresh = function () {
		$scope.refresh = true;
		$scope.GetPreviewData();
	};

	$scope.GoNext = function () {
		$scope.refresh = false;
		if ($scope.pageIndex < $scope.pageNumber)
			$scope.pageIndex += 1;

		$scope.GetPreviewData();
	};

	$scope.GoPrev = function () {
		$scope.refresh = false;
		if ($scope.pageIndex > 1)
			$scope.pageIndex -= 1;

		$scope.GetPreviewData();
	};

	$scope.GoFirst = function () {
		$scope.refresh = false;
		$scope.pageIndex = 1;

		$scope.GetPreviewData();
	};

	$scope.GoLast = function () {
		$scope.refresh = false;
		$scope.pageIndex = $scope.pageNumber;

		$scope.GetPreviewData();
	};

	$scope.SetPreviewPageImg = function (data) {
		var imgData = 'data:image/jpeg;base64,' + data.Text;
		$scope.pageNumber = data.FileCount;
		$('#preview_container > img')[0].src = imgData;
		$('#page_count').text($scope.pageIndex + ' of ' + $scope.pageNumber)
	};

	$scope.GetPreviewData = function () {
		$('#page-loading').show();
		$('#loader').show();
		var htmldata = $scope.trum.trumbowyg('html');
		$http({
			method: 'POST',
			url: asposeEditorAPI + 'GetPreviewData',
			data: {
				'htmldata': htmldata,
				'folderName': folderName,
				'fileName': fileName,
				'refresh': $scope.refresh,
				'pageIndex': $scope.pageIndex
			},
			responseType: "application/json"
		}).then(function (response) {
			var data = response.data;
			$scope.SetPreviewPageImg(data);
		}, function (error) {
			$scope.ShowError(error.data);
		}).finally(function () {
			$('#page-loading').fadeOut(600);
		});
	};

	$scope.Download = function (outputType) {

		$('#page-loading').show();
		$('#loader').show();

			$http({
				method: 'POST',
				url: asposeEditorAPI + 'UpdateContentsWithAttachments',
				data: {
					'createNewFile': createNewFile,
					'fileName': fileName,
					'folderName': folderName,
					'htmldata': $scope.trum.trumbowyg('html'),
					'outputType': outputType,
					'attachmentsData': attachmentsDict
				},
				responseType: "application/json"
			}).then(function (response) {
				var data = response.data;
				if (data.StatusCode === 200) {
					var fileLink = fileDownloadLink + "/" + data.FolderName + "?file=" + data.FileName;
					window.lastDownloadLink = fileLink;
					window.location = fileLink;
				}
				else
					$scope.ShowError(data.Status);
			}, function (error) {
				$scope.ShowError(error.data);
			}).finally(function () {
				$('#page-loading').fadeOut(600);
			});
	};

	$scope.ShowError = function (message) {
		$('#alert > p')[0].innerText = message;
		$('#alert').show();
	};

	$scope.Init();

});