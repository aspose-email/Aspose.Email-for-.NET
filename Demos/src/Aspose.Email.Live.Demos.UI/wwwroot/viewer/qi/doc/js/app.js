(function () {
	'use strict';
	var app = angular.module('myApp', [
		'ngSanitize',
		'ngAnimate',
		'ngQuantum',
		'ngResource'
	]);

	app.value('PageList', customPageList);

	app.run(['$templateCache', '$cacheFactory',
		function ($templateCache, $cacheFactory) {
			$templateCache = false;
		}]);

	app.config(['$httpProvider',
		function ($httpProvider) {
			$httpProvider.defaults.cache = false;
		}]);

	app.config(['$rootScopeProvider',
		function ($rootScopeProvider) {
			$rootScopeProvider.digestTtl(5);
		}]);

	app.directive('infinityscroll', function () {
		return {
			restrict: 'A',
			link: function (scope, element, attrs) {
				element.bind('scroll', function () {
					if ((element[0].scrollTop + element[0].offsetHeight) >= element[0].scrollHeight) {
						//scroll reach to end
						scope.$apply(attrs.infinityscroll);
					}
				});
			}
		}
	});

	app.directive('myEnter', function () {
		return function ($scope, element, attrs) {
			element.bind("keydown keypress", function (event) {
				if (event.which === 13) {
					$scope.$apply(function () {
						$scope.$eval(attrs.myEnter);
					});
					event.preventDefault();
				}
			});
		};
	});

	app.factory('apiService', function ($http) {
		var getData = function () {
			return $http({ method: "GET", url: apiURL + 'api/AsposeEmailViewer/DocumentPages?' + 'file=' + fileName + '&folderName=' + folderName }).then(function (result) {
				return result;
			});
		};

		return { getData: getData };
	});

	function GetNodeIndex(arr, node) {
        for (var j = 0; j < arr.length; j++) {
			if (CompareImageNode(arr[j], node)) {
                return j;
            }
		}

        return -1;
    }

    function CompareImageNode(node1, node2) {
		return (node1.ImageName === node2.ImageName);
	}

	function ProcessStringListResponse(value, $scope) {
        if (value.toLowerCase().includes('c:')) {
	        if (!customPageList.includes(value)) {
		        customPageList.push(value);
	        }
        }
    }

	function ProcessJsonResponse(value, $scope) {
		if (isNaN(value)) {
			var item;
            try {
				item = JSON.parse(value);
			} catch (e) {
				return ProcessStringListResponse(value, $scope);
            }

			item.ImageWidth = 640;
			item.ImageHeight = 480;
            if (item.ImageSize) {
	            var dimensions = item.ImageSize.split('x');
				if (dimensions.length === 2) {
					item.ImageWidth = Number(dimensions[0]);
					item.ImageHeight = Number(dimensions[1]);
				}
			}

            var nodeIdx = GetNodeIndex(customPageList, item);
			if (nodeIdx < 0) {
				nodeIdx = customPageList.length;
				customPageList.push(item);
			}

			SetPageSize(nodeIdx, item.ImageWidth, item.ImageHeight);
        }
    }

	function GetPagesData($scope, apiService) {


		var myDataPromise = apiService.getData();

		$scope.loading.show();
		myDataPromise.then(function (result) {
			var i = 0;

			angular.forEach(result.data, function (value) {
				if (i === 0 && result.data.length > 1) {
					totalPages = parseInt(value);
					UpdatePager();
					i++;
				}
				else {
					ProcessJsonResponse(value, $scope);

					if (customPageList.length > 0 && currentPageCount === 1) {
						var dvPages = document.getElementsByName("dvPages")[0];
						dvPages.style.cssText = "height: 100vh; padding-top: 55px; width: auto!important; overflow: auto!important; background-color: #777; background-image: none!important;";
						dvPages.getElementsByClassName('container-fluid')[0].classList.remove('hidden');
						// $scope.navigatePage('+');
					}

					i++;
				}

				$scope.loading.hide();
			});

			/*if (currentPageCount < totalPages) {
				$scope.NextPage();
			} */
		});
	}

	app.controller('ViewerAPIController',
		function ViewerAPIController($scope, $sce, $http, $window, apiService, $loading, $timeout, $q, $alert) {
			var $that = this;

			$scope.loadingButtonSucces = function () {
				return $timeout(function () {
					return true;
				}, 2000)
			}
			$scope.exitApp = function () {
				if (window.parent && window.parent.closeIframe) {
					window.history.back();
					window.parent.closeIframe();
				}
				else {

					if (callbackURL !== '')
						window.location = callbackURL + '?folderName=' + folderName + '&fileName=' + fileName;
					else if (featureName !== '')
						window.location = '/email/viewer/' + featureName;
					else
						window.location = '/email/viewer';
				}
			}
			$scope.loading = new $loading({
				busyText: ' Please wait while page loading...',
				theme: 'info',
				timeout: false,
				showSpinner: true
			});

			$scope.print = function() {
				window.print();
			}

			$scope.getError = function () {
				var deferred = $q.defer();

				setTimeout(function () {
					deferred.reject('Error');
				}, 1000);
				return deferred.promise;
			}

			$scope.displayAlert = function (title, message, theme) {
				$alert(message, title, theme)
			}

			if (customPageList.length <= 0) {
				$scope.PageList = customPageList;
			}

			GetPagesData($scope, apiService);

			/*$scope.NextPage = function () {

				if (currentPageCount > totalPages) {
					currentPageCount = totalPages;
					return;
				}

				if (currentPageCount <= totalPages) {
					currentPageCount += 1;
					currentSelectedPage = currentPageCount;

					if ($scope.PageList.length < currentPageCount) {
						GetPageData($scope, apiService, currentPageCount);
						currentSelectedPage = currentPageCount - 2;
					}
				}
			}*/

			$scope.selected = false;

			$scope.slectedPageImage = function (event, pageData) {
				var domId = event.target.id.replace('img-page-', '').replace('imgt-page-', '');
				currentSelectedPage = parseInt(event.target.id.replace('img-page-', '').replace('imgt-page-', ''));

				UpdatePager();

				if (event.target.id.startsWith('imgt-page-')) {
					location.hash = 'page-view-' + domId;
					$scope.selected = pageData;
				}
			}

			$scope.navigatePage = function (options) {
				if (options === '+') {
					currentPageCount += 1;
					if (currentPageCount > totalPages) {
						currentPageCount = totalPages;
					}
				}
				else if (options === '-') {
					currentPageCount -= 1;

					if (currentPageCount < 1) {
						currentPageCount = 1;
					}
				}
				else if (options === 'f') {
					currentPageCount = 1;
				}
				else if (options === 'e') {
					currentPageCount = totalPages;
				}
				else {
					if (document.getElementById('inputcurrentpage').value !== '')
						currentPageCount = parseInt(document.getElementById('inputcurrentpage').value);

					if (currentPageCount > totalPages) {
						currentPageCount = totalPages;
					}

					if (currentPageCount < 1) {
						currentPageCount = 1;
					}
				}

				currentSelectedPage = currentPageCount;

				if ($scope.PageList.length < currentSelectedPage) {
					//GetPageData($scope, apiService, currentPageCount);
					$scope.$broadcast('UpdatePages');
					$scope.$broadcast('UpdateThumbnails');
				}

				UpdatePager();
				location.hash = 'page-view-' + currentSelectedPage;
			};

			$scope.createPageImage = function (indx, pageData) {
				if (!pageData) {
					return pageData;
				}

				if (indx <= (imagedata.length - 1)) {
					return imagedata[indx];
				}
				else {
					prevoiusIndx = indx;
					var imgData = $sce.trustAsResourceUrl(apiURL + 'api/AsposeEmailViewer/pageimage?imageFolderName=' + encodeURIComponent(pageData.ImageFolderName) + '&imageFileName=' + encodeURIComponent(pageData.ImageName));
					imagedata.push(imgData);
					return imagedata[indx];
				}
			}

			$scope.itemSelected = '1.00';

			$scope.zoomPage = function (zoomOption) {
				$scope.itemSelected = zoomOption;
				ZoomPages(zoomOption);
			}

			$scope.getPageCss = function(pageIndex) {
				return GetPageCss(pageIndex);
			}
		}
	);
})();