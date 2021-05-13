using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Linq;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    using EmailApps = Aspose.Email.Live.Demos.UI.Models.EmailApps;
    public class EmailController : BaseController
	{
		public const int DefaultCacheDuration = 60 * 60 * 12;
		public IMemoryCache Cache { get; }

		public EmailController(IMemoryCache cache)
		{
            Cache = cache;
        }

		public string Locale
		{
			get
			{
				var routeValues = RouteData.Values;
				if (routeValues.ContainsKey("locale") && !string.IsNullOrEmpty(routeValues["locale"].ToString()))
				{
					return routeValues["locale"].ToString().ToUpper();
				}
				else
				{
					return "EN";
				}
			}
		}

        private IActionResult TryOpenEmailApp(EmailApps app, Action<ViewModel> setupAction = null) =>
			TryOpenEmailApp(app, null, setupAction);

		private IActionResult TryOpenEmailApp(EmailApps app, string extension, Action<ViewModel> setupAction = null) =>
			TryOpenEmailApp(app, app.ToString(), extension, setupAction);

		private IActionResult TryOpenEmailApp(EmailApps app, string appName, string extension, Action<ViewModel> setupAction = null)
		{
			var uri = RequestUrl;
			var url = uri.OriginalString.ToLower();
			var routeValues = RouteData.Values;

			// ---- Prepare resources
			var locale = Locale;
			var appResourcesKey = $"appresources_{appName}_{locale}";

			var appResources = Cache.GetOrCreate<IAppResources>(appResourcesKey, entry =>
			{
				entry.Priority = CacheItemPriority.NeverRemove;
				return new PrefetchedAppResources(Resources, uri, appName);
			});

			// ---- Setup a model
			var model = new ViewModel(appResources, this, uri, app, locale, appName, extension);
			setupAction?.Invoke(model);

			model.OnBeforeRendering();

			return View(model);
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Conversion(string extension)
		{
			return TryOpenEmailApp(EmailApps.Conversion, extension, model =>
			{
				model.SaveAsComponent = true;
				model.SaveAsOriginal = false;
				model.SaveAsOptions = new string[]
				{
					"eml",
					"msg",
					"pst",
					"mbox",
					"html",
					"mht",
					"jpg",
					"png",
					"svg",
					"bmp",
					"tiff",
					"pdf",
					"doc",
					"ppt",
					"rtf",
					"docx",
					"docm",
					"dotx",
					"dotm",
					"odt",
					"ott",
					"epub",
					"txt",
					"emf",
					"xps",
					"pcl",
					"ps",
					"mhtml"
				};
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Viewer()
		{
			return TryOpenEmailApp(EmailApps.Viewer, model =>
			{
				model.UploadAndRedirect = true;
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Editor()
		{
			return TryOpenEmailApp(EmailApps.Editor, model =>
			{
				model.CanWorkWithoutFiles = true;
				model.UploadAndRedirect = true;
				model.NeedsProcessButton = true;
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Parser()
		{
			return TryOpenEmailApp(EmailApps.Parser);
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Metadata()
		{
			return TryOpenEmailApp(EmailApps.Metadata, model =>
			{
				model.NeedsDownloadForm = true;
				model.UploadAndRedirect = true;
				model.ControlsView = "MetadataControls";
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Merger()
		{
			return TryOpenEmailApp(EmailApps.Merger, model =>
			{
				model.UseSorting = true;
				model.SaveAsComponent = false;
				model.SaveAsOriginal = false;
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Redaction()
		{
			return TryOpenEmailApp(EmailApps.Redaction, model =>
			{
				model.ControlsView = "RedactionControls";
				model.SaveAsComponent = false;
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Search()
		{
			return TryOpenEmailApp(EmailApps.Search, model =>
			{
				model.ControlsView = "SearchControls";
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Headers()
		{
			return TryOpenEmailApp(EmailApps.Headers, model =>
			{
				model.CanWorkWithoutFiles = true;
				model.ControlsView = "HeadersControls";
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Signature()
		{
			return TryOpenEmailApp(EmailApps.Signature, model =>
			{
				model.ControlsView = "SignatureControls";
				model.SaveAsComponent = true;
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Annotation()
		{
			return TryOpenEmailApp(EmailApps.Annotation);
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Comparison()
		{
			return TryOpenEmailApp(EmailApps.Comparison);
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Watermark()
		{
			return TryOpenEmailApp(EmailApps.Watermark, model =>
			{
				model.ControlsView = "WatermarkControls";
			});
		}

		[ResponseCache(Location = ResponseCacheLocation.Any, Duration = DefaultCacheDuration)]
		public IActionResult Assembly()
		{
			return TryOpenEmailApp(EmailApps.Assembly, model =>
			{
				model.DefaultFileBlockDisabled = true;
				model.ControlsView = "AssemblyControls";
				model.NeedsDownloadForm = true;
				model.UploadAndRedirect = true;
			});
		}
	}
}
