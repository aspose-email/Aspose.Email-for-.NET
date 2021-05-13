namespace Aspose.Email.Live.Demos.UI.Models
{
    ///<Summary>
    /// License class to set apose products license
    ///</Summary>
    public static class License
	{
		const string TotalNetLicense = "Aspose.Total.NET.lic";

		///<Summary>
		/// SetAsposeWordsLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeWordsLicense()
		{
			Aspose.Words.License awLic = new Aspose.Words.License();
			awLic.SetLicense(TotalNetLicense);
		}
		///<Summary>
		/// SetAsposeCellsLicense method to Aspose.Cells License
		///</Summary>
		public static void SetAsposeCellsLicense()
		{
			Aspose.Cells.License acLic = new Aspose.Cells.License();
			acLic.SetLicense(TotalNetLicense);
		}
		///<Summary>
		/// SetAsposeEmailLicense method to Aspose.Email License
		///</Summary>
		public static void SetAsposeEmailLicense()
		{
			Aspose.Email.License acLic = new Aspose.Email.License();
			acLic.SetLicense(TotalNetLicense);
		}
		///<Summary>
		/// SetAsposeSlidesLicense method to Aspose.Slides License
		///</Summary>
		public static void SetAsposeSlidesLicense()
		{
			Aspose.Slides.License acLic = new Aspose.Slides.License();
			acLic.SetLicense(TotalNetLicense);
		}

		///<Summary>
		/// SetAsposeImagingLicense method to Aspose.Imaging License
		///</Summary>
		public static void SetAsposeImagingLicense()
		{
			Aspose.Imaging.License lic = new Aspose.Imaging.License();
			lic.SetLicense(TotalNetLicense);
		}
		///<Summary>
		/// SetAsposeHtmlLicense method to Aspose.Html License
		///</Summary>
		public static void SetAsposeHtmlLicense()
		{
			Aspose.Html.License lic = new Aspose.Html.License();
			lic.SetLicense(TotalNetLicense);
		}
	}
}
