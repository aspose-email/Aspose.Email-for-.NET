using Aspose.Email.Live.Demos.UI.UI.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace Aspose.Email.Live.Demos.UI.UI.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
