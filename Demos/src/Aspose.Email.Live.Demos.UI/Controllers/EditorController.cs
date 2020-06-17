using Aspose.Email.Live.Demos.UI.Models.Common;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class EditorController : BaseController
	{
		public override string Product => (string)RouteData.Values["product"];

		public ActionResult Editor()
		{
			var model = new ViewModel(this, "Editor")
			{				
				MaximumUploadFiles = 1,				
				DropOrUploadFileLabel = Resources["DropOrUploadFile"],
				UploadAndRedirect = true
			};
			if (model.RedirectToMainApp)
				return Redirect("/email/" + model.AppName.ToLower());
			return View(model);			
		}
	}
}
