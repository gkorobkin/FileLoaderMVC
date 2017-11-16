using System;
using System.Diagnostics;
using System.Web;
using System.Web.Mvc;
using static System.IO.Path;

namespace FileLoaderMVC.Controllers
{
    public class BatchPreTradeController : Controller
    {
        private IBatchProcessor _processor;

        public BatchPreTradeController()
        {
            _processor = new BatchPreTradeProcessor();
        }

        // GET: BatchPreTrade
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult UploadBatch()
        {
            return View();
        }

        // POST: BatchPreTrade
        [HttpPost]
        public ActionResult UploadBatch(HttpPostedFileBase file)
        {
            try
            {
                if (file?.ContentLength > 0)
                {
                    var fil = GetFileName(file?.FileName);
                    if (!string.IsNullOrWhiteSpace(fil))
                    {
                        var path = Combine(Server.MapPath("~/UploadedFiles"), fil);
                        file.SaveAs(path);

                        _processor?.Process(path);
                    }
                }
                ViewBag.Message = "Pretrade Batch uploaded Successfully!";
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                ViewBag.Message = "Pretrade Batch upload failed!";
            }

            return View();
        }
    }
}