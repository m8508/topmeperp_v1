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
        InquiryFormService service = new InquiryFormService();
        // GET: Inquiry
        public ActionResult Index(string id)
        {
            log.Info("inquiry index : projectid=" + id);
            ViewBag.projectId = id;
            return View();
        }
        // POST : Search
        [HttpPost]
        public ActionResult Index(FormCollection f)
        {
            TnderProject s = new TnderProject();
            log.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            List<topmeperp.Models.TND_PROJECT_ITEM> lstProject = s.getProjectItem(Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            return View("Index", lstProject);
        }
        //Create Project Form
        [HttpPost]
        public ActionResult Create(TND_PROJECT_FORM qf)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["prjId"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["prjId"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            //建立空白詢價單
            log.Info("create new form template");
            TnderProject s = new TnderProject();
            SYS_USER u = (SYS_USER)Session["user"];
            qf.PROJECT_ID = Request["prjId"];
            qf.CREATE_ID = u.USER_ID;
            qf.CREATE_DATE = DateTime.Now;
            TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
            string fid = s.newForm(qf, lstItemId);
            return RedirectToAction("InquiryMainPage/" + qf.PROJECT_ID);
        }
        //測試詢價單下載
        public ActionResult ExportInquiry()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ExportInquiry(FormCollection form)
        {
            log.Info("get inquiry form:formid=" + form["formid"]);
            service.getInqueryForm(form["formid"]);
            InquiryFormToExcel poi = new InquiryFormToExcel();
            poi.exportExcel(service.formInquiry, service.formInquiryItem);
            return View();
        }

        public ActionResult InquiryMainPage(string id)
        {
            log.Info("queryInquiry by projectID=" + Request["projectid"]);
            List<TND_PROJECT_FORM> lst = null;
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                lst = service.getFormTemplateByProject(id);
            }
            return View(lst);
        }

        //上傳廠商報價單
        public string FileUpload(HttpPostedFileBase file)
        {
            log.Info("Upload form from supplier:"+ Request["projectid"]);
            string projectid = Request["projectid"];
            //上傳至廠商報價單目錄
            if (null!=file && file.ContentLength != 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid+"/"+ ContextService.quotesFolder, fileName);
                file.SaveAs(path);
                log.Info("Parser Excel File Begin:" + file.FileName);
                InquiryFormToExcel quoteFormService = new InquiryFormToExcel();
                quoteFormService.convertInquiry2Project(path, projectid);
                int i =service.createInquiryFormFromSupplier(quoteFormService.form, quoteFormService.formItems);
                log.Info("add supplier form record count=" + i);
            }
            return "檔案匯入成功!!";
        }
    }
}



