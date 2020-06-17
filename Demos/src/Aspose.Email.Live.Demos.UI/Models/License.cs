using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Email.Live.Demos.UI.Models
{
	///<Summary>
	/// License class to set apose products license
	///</Summary>
	public static class License
	{
		private static string _licenseFileName = "Aspose.Total.lic";


		///<Summary>
		/// SetAsposeEmailLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeEmailLicense()
		{
			try
			{
				Aspose.Email.License acLic = new Aspose.Email.License();
				acLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		///<Summary>
		/// SetAsposeSlidesLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeSlidesLicense()
		{
			try
			{
				Aspose.Slides.License acLic = new Aspose.Slides.License();
				acLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		///<Summary>
		/// SetAsposeWordsLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeWordsLicense()
		{
			try
			{
				Aspose.Words.License acLic = new Aspose.Words.License();
				acLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
		///<Summary>
		/// SetAsposeImagingLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeImagingLicense()
		{
			try
			{
				Aspose.Imaging.License acLic = new Aspose.Imaging.License();
				acLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}


	}
}
