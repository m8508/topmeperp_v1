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
        public ActionResult Index(string id)
        {
            log.Info("inquiry index : projectid="+ id);
            ViewBag.projectId = id;
            return View();
        }

        // POST : Search
        [HttpPost ]
        public ActionResult Index(FormCollection f)
        {
            TnderProject s = new TnderProject();
            log.Info("projectid="+ Request["projectid"]+",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            List <topmeperp.Models.TND_PROJECT_ITEM> lstProject = s.getProjectItem(Request["projectid"],Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            return View("Index", lstProject);
        }
        //Create Project Form
        public ActionResult Create()
        {
            log.Info("item_list:" + Request["chkItem"]);
            //TnderProject service = new TnderProject();
            return View("~/Views/Inquiry/FormTemplate.cshtml");
        }
      }
    }


