using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;

namespace topmeperp.Controllers
{
    public class InquiryController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));

        // GET: Inquiry
        public ActionResult Index()
        {
            log.Info("inquiry index!");
            return View();
        }

        // POST : Search
        public ActionResult Search()
        {
            TnderProject s = new TnderProject();
            log.Info("textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            List <topmeperp.Models.TND_PROJECT_ITEM> lstProject = s.getProjectItem(Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View("Index", lstProject);
        }


    }
}

