using Aspose.Email.Live.Demos.UI.Controllers;
using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models.Email;
using Aspose.Email.Live.Demos.UI.Services;
using System;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Models
{
    public partial class ViewModel
	{
		public const int MaximumUploadFiles = 10;
		public BaseController Controller { get; }

		public Uri RequestUri {get;}
		public EmailApps AppType { get; }

		public string Product => "email";
		public string ProductAppName { get; }

		public string DocumentNaming => this["Document"];

		public string PageProductTitle => this["PageProductTitle", "Aspose.Email"];

		public string AppName { get; }

		public string AppURL { get; }
		public string AppRoute { get; }

		/// <summary>
		/// File extension without dot received by "fileformat" value in RouteData (e.g. docx)
		/// </summary>
		public string Extension1 { get; }

		/// <summary>
		/// File extension without dot received by "fileformat" value in RouteData (e.g. docx)
		/// </summary>
		public string Extension2 { get; }

		public bool IsCanonical => string.IsNullOrEmpty(Extension1) && string.IsNullOrEmpty(Extension2);

		/// <summary>
		/// Name of the partial View of controls (e.g. SignatureControls)
		/// </summary>
		public string ControlsView { get; set; }

		public string AnotherFileText { get; set; }
		public string UploadButtonText { get; set; }
		public string ViewerButtonText { get; set; }
		public bool ShowViewerButton { get; set; }
		public string SuccessMessage { get; set; }

		/// <summary>
		/// List of app features for ul-list. E.g. Resources[app + "LiFeature1"]
		/// </summary>
		public List<string> AppFeatures { get; set; }

		public string Title { get; set; }
		public string TitleSub { get; set; }

		public string PageTitle
		{
			get => Controller.ViewBag.PageTitle;
			set => Controller.ViewBag.PageTitle = value;
		}

		public string MetaDescription
		{
			get => Controller.ViewBag.MetaDescription;
			set => Controller.ViewBag.MetaDescription = value;
		}

		public string MetaKeywords
		{
			get => Controller.ViewBag.MetaKeywords;
			set => Controller.ViewBag.MetaKeywords = value;
		}

		/// <summary>
		/// If the application doesn't need to upload several files (e.g. Viewer, Editor). Start the processing instantly after the dropping a file
		/// </summary>
		public bool UploadAndRedirect { get; set; }


		/// <summary>
		/// If the application needs custom process button when enabled <see cref="UploadAndRedirect"/> option
		/// </summary>
		public bool NeedsProcessButton { get; set; }

		/// <summary>
		/// If the application needs download form when enabled <see cref="UploadAndRedirect"/> option
		/// </summary>
		public bool NeedsDownloadForm { get; set; }

		protected string TitleCase(string value) => new System.Globalization.CultureInfo("en-US", false).TextInfo.ToTitleCase(value);

		/// <summary>
		/// e.g., .doc|.docx|.dot|.dotx|.rtf|.odt|.ott|.txt|.html|.xhtml|.mhtml
		/// </summary>
		public string ExtensionsString { get; set; }

		#region SaveAs
		protected bool _saveAsComponent;
		public bool SaveAsComponent
		{
			get => _saveAsComponent;
			set
			{
				_saveAsComponent = value;
				Controller.ViewBag.SaveAsComponent = value;

				if (_saveAsComponent)
				{
					SaveAsOptions = this["SaveAsOptions", ""].Split(',');

					if (AppFeatures != null)
					{
						if (AppResources.ContainsResource("SaveAsLiFeature"))
							AppFeatures.Add(this["SaveAsLiFeature"]);
					}
				}
			}
		}

		/// <summary>
		/// FileFormats in UpperCase
		/// </summary>
		public string[] SaveAsOptions { get; set; }

		public Dictionary<string, string[]> SaveAsOptionsSpecific { get; set; }

		/// <summary>
		/// Original file format SaveAs option for multiple files uploading
		/// </summary>
		public bool SaveAsOriginal { get; set; }
		#endregion

		/// <summary>
		/// The possibility of changing the order of uploaded files. It is actual for Merger App.
		/// </summary>
		public bool UseSorting { get; set; }

		public bool ShowFileDropButtonBar { get; set; }

		public bool HowToPanelEnabled { get; set; }
		public bool ShowHowTo => HowToModel != null && HowToPanelEnabled;
		public EmailHowToModel HowToModel { get; set; }

		public bool FAQPagePanelEnabled { get; set; }
		public bool ShowFAQPage => FAQPageModel != null && FAQPageModel.List.Count > 0 && FAQPagePanelEnabled;
		public EmailFAQPageModel FAQPageModel { get; set; }

		public IEnumerable<EmailAnotherApp> OtherApps { get; set; }

		public string JSOptions { get; private set; }

		/// <summary>
		/// The possibility of main button working with no files uploaded
		/// </summary>
		public bool CanWorkWithoutFiles { get; set; }

		public bool DefaultFileBlockDisabled { get; set; }
		public IAppResources AppResources { get; }


		public string PopularFeaturesTitle { get; private set; }
		public string PopularFeaturesTitleSub { get; private set; }
		public string OtherFeaturesTitle { get; private set; }
		public string OtherFeaturesTitleSub { get; private set; }

		public string VideoUrl { get; private set; }
		public string VideoTitle { get; private set; }
		public string Locale { get; private set; }

		public string BaseUrl
		{
			get
			{
				var request = Controller.Request;
				return $"{request.Scheme}://{request.Host}{request.PathBase}";
			}
		}

		public bool IsDummyPage { get; set; }

		public bool OverviewPanelEnabled { get; set; }
		public bool AppFeaturesPanelEnabled { get; set; }
		public bool AppFeaturesActionStringEnabled { get; set; }


		public string AppFeaturesTitle => this["AppFeaturesTitle", "Aspose.Email " + AppName];

		/// <summary>
		/// Product, (AppName, AnotherApp)
		/// </summary>
		protected Dictionary<string, Dictionary<string, EmailAnotherApp>> OtherAppsStatic = new Dictionary<string, Dictionary<string, EmailAnotherApp>>();

		public ViewModel(IAppResources resources, BaseController controller, Uri requestUri, EmailApps appType, string locale, string app, string extension1 = null, string extension2 = null)
		{
			Locale = locale;
			Controller = controller;
			Extension1 = extension1;
			Extension2 = extension2;
			RequestUri = requestUri;

			AppResources = resources; 

			AppName = this[$"APPName", app];
			ProductAppName = Product + app;
			AppType = appType;

			var url = requestUri.OriginalString;
			AppURL = url.Substring(0, (url.IndexOf("?") > 0 ? url.IndexOf("?") : url.Length));
			AppRoute = requestUri.GetLeftPart(UriPartial.Path);

			InitOtherApps(Product.ToLower(), AppURL);

			PrepareHowToModel();

			SetTitles();
			SetAppFeatures(app);
			PrepareFAQPageModel();

			if (OtherAppsStatic.ContainsKey(Product.ToLower()))
				OtherApps = OtherAppsStatic[Product.ToLower()].Values;

			ShowViewerButton = true;
			SaveAsOriginal = true;
			SaveAsComponent = false;
		}

		protected void PrepareHowToModel()
		{
			if (!string.IsNullOrEmpty(Extension1) || IsCanonical)
			{
				try
				{
					HowToModel = new EmailHowToModel(this);
				}
				catch
				{
					HowToModel = null;
				}
			}
		}

		protected void PrepareFAQPageModel()
		{
			FAQPageModel = new EmailFAQPageModel(this);
		}

		/// <summary>
		/// Returns Word, OpenOffice, RTF or Text
		/// </summary>
		/// <param name="extension"></param>
		/// <param name="defaultValue"></param>
		/// <returns></returns>
		public string DesktopAppNameByExtension(string extension, string defaultValue = null)
		{
			if (!string.IsNullOrEmpty(extension))
			{
				switch (extension.ToLower())
				{
					case "docx":
					case "doc":
					case "dot":
					case "dotx":
						return "Word";
					case "odt":
					case "ott":
						return "OpenOffice";
					case "rtf":
						return "RTF";
					case "txt":
						return this["Text"];
					case "md":
						return "Markdown";
					case "ps":
						return "PostScript";
					case "tex":
						return "LaTeX";
					case "acroform":
						return this["pdfXfaToAcroform"];
					case "pdfa1a":
						return "PDF/A-1A";
					case "pdfa1b":
						return "PDF/A-1B";
					case "pdfa2a":
						return "PDF/A-2A";
					case "pdfa3a":
						return "PDF/A-3A";
					default:
						return string.IsNullOrEmpty(defaultValue) ? extension.ToUpper() : defaultValue;
				}
			}

			return defaultValue;
		}

		void SetTitles()
		{
			Title = this["H1"];
			TitleSub = this["H4"];

			VideoUrl = this["DefaultVideoUrl", null];
			VideoTitle = this["DefaultVideoTitle", null];

			PageTitle = this["PageTitle"];
			MetaDescription = this["MetaDescription"];
			MetaKeywords = this["MetaKeywords"];
			ExtensionsString = this["ValidationExpression"];

			OverviewPanelEnabled = this["OverviewPanelEnabled"].ParseToBoolOrDefault();
			AppFeaturesPanelEnabled = this["AppFeaturesPanelEnabled"].ParseToBoolOrDefault();
			
			FAQPagePanelEnabled = this["FAQPagePanelEnabled"].ParseToBoolOrDefault();
			HowToPanelEnabled = this["HowToPanelEnabled"].ParseToBoolOrDefault();
			AppFeaturesActionStringEnabled = this["AppFeaturesActionStringEnabled"].ParseToBoolOrDefault();
		}

		void SetAppFeatures(string app)
		{
			AppFeatures = new List<string>();

			var i = 1;

			string value;
			while ((value = AppResources.GetResourceOrDefault($"LiFeature{i++}")) != null)
				AppFeatures.Add(ResourceHelper.InjectFormatLinks(value));

			// Stop other developers to add unnecessary features.
			if (AppFeatures.Count == 0)
			{
				i = 1;

				while ((value = AppResources.GetResourceOrDefault($"LiFeature{i}")) != null)
				{
					if (!value.Contains("Instantly download") || AppFeatures.Count == 0)
						AppFeatures.Add(ResourceHelper.InjectFormatLinks(value));
					i++;
				}
			}
		}

	    void InitOtherApps(string product, string appUrl = null)
		{
			var appList = new Dictionary<string, EmailAnotherApp>();

			var apps = new[]
			{
				"Conversion",
				"Viewer",
				"Headers",
				"Metadata",
				"Parser",
				"Search",
				"Merger",
				"Signature",
				"Watermark",
				"Assembly",
				"Redaction",
				"Annotation",
				"Editor",
				"Comparison"
			};

			foreach (var appName in apps)
				appList.Add(appName, new EmailAnotherApp(appName, Locale));

			OtherAppsStatic[product] = appList;
		}

		public string this[string key] => AppResources.GetResourceOrDefault(key);

		public string this[string key, string defaultValue]
		{
			get
			{
				var value = AppResources.GetResourceOrDefault(key);

				if (value != null)
				{
					return value;
				}

				return defaultValue;
			}
		}

		public void OnBeforeRendering()
		{
			JSOptions = new JSOptions(this).ToString();

			// if the app has default process button
			if (!UploadAndRedirect || NeedsProcessButton)
				UploadButtonText = this["Button"];

			// if the app has result downloading section
			if (!UploadAndRedirect || NeedsDownloadForm)
			{
				ViewerButtonText = this["Viewer", "VIEW RESULTS"];
				SuccessMessage = this["SuccessMessage", "Your file has been processed successfully"];
				AnotherFileText = this["AnotherFile", "Process another file"];
			}
		}
	}
}
