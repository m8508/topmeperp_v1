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
            //取得專案編號
            log.Info("Project Id:" + Request["prjId"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["prjId"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i=0;
            for (i = 0; i < lstItemId.Count(); i++) {
            log.Info("item_list return:" + lstItemId[i]); }
            //建立空白詢價單
            log.Info("create new form template");
            TnderProject s = new TnderProject();
            //PROJECT_FORM_ITEM 可由lstItemId取得對應的標單編號(PROJECT_ITEM)
            //List<topmeperp.Models.TND_PROJECT_ITEM> lstProjectItem = s.getProjectItemId(Request["prjId"], Request["chkItem"]);
            // return View("Create", lstProjectItem);
            //發現問題先註解掉兩行(上面)
            return View();
        }
        [HttpPost]
        public ActionResult ProjectForm(TND_PROJECT_FORM qf)
        //PROJECT_FOM 完全是需要新增的
        {
            log.Info("create inquiry form process! project id=" + qf.PROJECT_ID);
            TnderProject service = new TnderProject();
            service.newForm(qf);
            return View("~/Views/Inquiry/Index.cshtml");
        }
    }
}
    


