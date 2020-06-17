using System;
using System.Data;
using System.Configuration;
using System.Web;

/// <summary>
/// Summary description for Configuration
/// </summary>
namespace Aspose.Email.Live.Demos.UI.Config
{
	public static class Configuration
	{		
		private static string _appName = ConfigurationManager.AppSettings["AppName"].ToString();
		private static string _asposeEmailLiveDemosPath = ConfigurationManager.AppSettings["AsposeEmailLiveDemosPath"].ToString();
		private static string _resourceFileSessionName = ConfigurationManager.AppSettings["ResourceFileSessionName"];	      
		private static string _fileViewLink = ConfigurationManager.AppSettings["FileViewLink"];		
		private static string _productsAsposeEmailAssetURL = ConfigurationManager.AppSettings["ProductsAsposeEmailAssetURL"];
		private static string _fileDownloadLink = ConfigurationManager.AppSettings["FileDownloadLink"];
		

		public static string ResourceFileSessionName
		{
			get { return _resourceFileSessionName; }
		}
		/// <summary>
		/// Get Working Directory
		/// </summary>
		public static string WorkingDirectory
		{
			
		get {
				string sourceFilespath = "";
				if (HttpContext.Current == null)
				{
					sourceFilespath = System.Web.Hosting.HostingEnvironment.MapPath("~/Assets/SourceFiles/");
				}
				else
				{

					sourceFilespath = HttpContext.Current.Server.MapPath("~/Assets/SourceFiles/");
				}
				if ( ! System.IO.Directory.Exists(sourceFilespath))
				{
					System.IO.Directory.CreateDirectory(sourceFilespath);
				}

				return sourceFilespath;
			}
		}

		/// <summary>
		/// Get Working Directory
		/// </summary>
		public static string OutputDirectory
		{
			get
			{
				string OutputFilespath = "";
				if (HttpContext.Current == null)
				{
					OutputFilespath = System.Web.Hosting.HostingEnvironment.MapPath("~/Assets/OutputFiles/");
				}
				else
				{

					OutputFilespath = HttpContext.Current.Server.MapPath("~/Assets/OutputFiles/");
				}

				
				if (!System.IO.Directory.Exists(OutputFilespath))
				{
					System.IO.Directory.CreateDirectory(OutputFilespath);
				}

				return OutputFilespath;
			}
		}		
		public static string ProductsAsposeEmailAssetURL
		{
			get { return _productsAsposeEmailAssetURL; }
		}	
		
		public static string AppName
        {
            get { return _appName; }
        }
        public static string AsposeEmailLiveDemosPath
		{
            get { return _asposeEmailLiveDemosPath; }
        }
		public static string FileDownloadLink
		{
			get { return _fileDownloadLink; }
		}
		
		public static string FileViewLink
		{
			get { return _fileViewLink; }
		}
	}
}
