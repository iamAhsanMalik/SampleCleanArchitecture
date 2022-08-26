using DHAAccounts.Models;
using System;
using System.IO;
using System.Net;
using System.Web.Mvc;

namespace Accounting.Controllers
{
    public class PostingsController : Controller
    {
        public ActionResult UploadFile()
        {
            string HostName = System.Web.HttpContext.Current.Request.Url.Host;
            ImportDepositFileViewModel model = new ImportDepositFileViewModel();
            string BaseFolder = @"C:\Users\STMDEV08\Desktop"; //live path
            foreach (string txtName in Directory.GetFiles(BaseFolder, "*.txt"))
            {
                string FileName = Path.GetFileName(txtName);
                Uri uri = new Uri(txtName.Replace("/download", ""));
                string myfilename = Path.GetFileName(uri.LocalPath);
                string savePath = @"/Users/STMDEV08/Desktop/files/" + SysPrefs.SubmissionFolder;
                if (!Directory.Exists(savePath)) Directory.CreateDirectory(savePath);
                savePath += "/" + FileName;
                Console.WriteLine("File 1 Path=> " + savePath);
                WebClient _webclient = new WebClient();
                _webclient.DownloadFile(txtName, savePath);
            }
            return View("~/Views/Home/UploadFile.cshtml", model);
        }
    }

}