using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using Aspose.Email.Live.Demos.UI.Config;

namespace Aspose.Email.Live.Demos.UI
{
	public class Global : HttpApplication
	{
		
		protected void Application_Error(object sender, EventArgs e)
		{			
			
		}

		void Application_Start(object sender, EventArgs e)
		{
			AreaRegistration.RegisterAllAreas();
			GlobalConfiguration.Configure(WebApiConfig.Register);
			RouteConfig.RegisterRoutes(RouteTable.Routes);			
			BundleConfig.RegisterBundles(BundleTable.Bundles);
			RegisterCustomRoutes(RouteTable.Routes);
		}
		void Session_Start(object sender, EventArgs e)
		{
			//Check URL to set language resource file
			string _language = "EN";
			
			SetResourceFile(_language);
		}

		private void SetResourceFile(string strLanguage)
		{
			if (Session["AsposeEmailResources"] == null)
				Session["AsposeEmailResources"] = new GlobalAppHelper(HttpContext.Current, Application, Configuration.ResourceFileSessionName, strLanguage);
		}
		
			void RegisterCustomRoutes(RouteCollection routes)
		{
			routes.RouteExistingFiles = true;
			routes.Ignore("{resource}.axd/{*pathInfo}");
					

			routes.MapRoute(
				name: "Default",
				url: "Default",
				defaults: new { controller = "Home", action = "Default" }
			);
			
			routes.MapRoute(
				"AsposeEmailConversionRoute",
				"{product}/Conversion",
				 new { controller = "Conversion", action = "Conversion" }
			);
			routes.MapRoute(
				"AsposeEmailRemoveAnnotationRoute",
				"annotation/remove",
				 new { controller = "Annotation", action = "Remove" }
			);
			routes.MapRoute(
				"AsposeEmailSignatureRoute",
				"{Product}/signature",
				 new { controller = "Signature", action = "Signature" }
			);
			routes.MapRoute(
				"AsposeEmailSearchRoute",
				"{product}/search",
				 new { controller = "Search", action = "Search" }
			);
			routes.MapRoute(
				"AsposeEmailEditorRoute",
				"{product}/editor",
				 new { controller = "Editor", action = "Editor" }
			);
			routes.MapPageRoute(
			  "AsposeEditorEmail",
			  "{product}/edit",
			  "~/Editor/Default.aspx"
			);
			routes.MapRoute(
				"AsposeEmailRedactionRoute",
				"{product}/redaction",
				 new { controller = "Redaction", action = "Redaction" }
			);
			routes.MapRoute(
				"AsposeEmailWatermarkRoute",
				"{product}/watermark",
				 new { controller = "Watermark", action = "Watermark" }
			);
			routes.MapRoute(
				"AsposeEmailHeadersRoute",
				"{product}/headers",
				 new { controller = "Headers", action = "Headers" }
			);
			routes.MapRoute(
				"AsposeEmailParserRoute",
				"{product}/parser",
				 new { controller = "Parser", action = "Parser" }
			);
			routes.MapRoute(
				"AsposeEmailAnnotationRoute",
				"{product}/annotation",
				 new { controller = "Annotation", action = "Annotation" }
			);
			routes.MapRoute(
				"AsposeEmailMetadataRoute",
				"{product}/metadata",
				 new { controller = "Metadata", action = "Metadata" }
			);
			routes.MapRoute(
				"AsposeEmailMergerRoute",
				"{product}/merger",
				 new { controller = "Merger", action = "Merger" }
			);
			routes.MapRoute(
				"AsposeEmailComparisonRoute",
				"{product}/comparison",
				 new { controller = "Comparison", action = "Comparison" }
			);
			routes.MapRoute(
				"AsposeEmailViewerRoute",
				"{product}/viewer",
				 new { controller = "Viewer", action = "Viewer" }
			);
			routes.MapPageRoute(
			  "AsposeEmailDefaultViewerRoute",
			  "email/view",
			  "~/ViewerApp/Default.aspx"
			);
			
			routes.MapRoute(
				"DownloadFileRoute",
				"common/download",
				new { controller = "Common", action = "DownloadFile" }				
				
			);
			routes.MapRoute(
				"UploadFileRoute",
				"common/uploadfile",
				new { controller = "Common", action = "UploadFile" }

			);
		}

		private void MapProductToolPageRoute(RouteCollection routes, string routeName, string routeUrl, string physicalFile, string productRegex)
		{
			routes.MapPageRoute(routeName, routeUrl, physicalFile, false, null, new RouteValueDictionary { { "Product", productRegex } });
		}
	}
}
