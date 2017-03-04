using NuGet;
using NuGet.Server;
using NuGet.Server.Publishing;
using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Pietastic.WebUpdateService.Controllers
{
    public class FilesController : Controller
    {
        // GET: Files
        public ActionResult Index()
        {
			var path = VirtualPathUtility.ToAbsolute("~/nuget");
			var url = new Uri(Request.Url, path).AbsoluteUri;
			IPackageRepository packageService = PackageRepositoryFactory.Default.CreateRepository(url);

			var packages = packageService.GetPackages().Where(q => q.IsLatestVersion && q.Listed).ToList();

			return View(packages);
        }

		public FileResult GetSoftware(string fileName)
		{
			return File(System.IO.File.ReadAllBytes(Server.MapPath("~/Content/Launcher/Pietastic.InstallerUpdater.exe")), "application/octet-stream", string.Format("{0}.exe", fileName));
		}
    }
}