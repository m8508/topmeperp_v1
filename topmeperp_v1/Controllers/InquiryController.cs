using log4net;
using System;
using System.Collections.Generic;
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
        public ActionResult Create()
        {
            return View();
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
            //產生詢價單實體檔案
            service.getInqueryForm(fid);
            InquiryFormToExcel poi = new InquiryFormToExcel();
            poi.exportExcel(service.formInquiry, service.formInquiryItem);
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
            InquiryFormModel formData = new InquiryFormModel();
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                formData.tndTemplateProjectForm = service.getFormTemplateByProject(id);
                formData.tndProjectFormFromSupplier = service.getFormByProject(id);
            }
            return View(formData);
        }

        //上傳廠商報價單
        public string FileUpload(HttpPostedFileBase file)
        {
            log.Info("Upload form from supplier:" + Request["projectid"]);
            string projectid = Request["projectid"];
            //上傳至廠商報價單目錄
            if (null != file && file.ContentLength != 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid + "/" + ContextService.quotesFolder, fileName);
                file.SaveAs(path);
                log.Info("Parser Excel File Begin:" + file.FileName);
                InquiryFormToExcel quoteFormService = new InquiryFormToExcel();
                quoteFormService.convertInquiry2Project(path, projectid);
                int i = service.createInquiryFormFromSupplier(quoteFormService.form, quoteFormService.formItems);
                log.Info("add supplier form record count=" + i);
            }
            return "檔案匯入成功!!";
        }
        //比價功能資料頁
        public ActionResult ComparisonMain(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);

            //取得專案基本資料
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.id = p.PROJECT_ID;
            ViewBag.projectName = p.PROJECT_NAME;

            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得主系統資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSystemMain(id))
            {
                log.Debug("Main System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            //取得次系統資料
            List<SelectListItem> selectSub = new List<SelectListItem>();
            foreach (string itm in service.getSystemSub(id))
            {
                log.Debug("Sub System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectSub.Add(selectI);
                }
            }
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;

            //設定查詢條件
            return View();
        }
        //取得比價資料
        [HttpPost]
        public ActionResult ComparisonData(FormCollection form)
        {
            //傳入查詢條件
            log.Info("start project id=" + form["id"] + ",TypeCode1=" + form["typeCode1"] + ",typecode2=" + form["typeCode2"] + ",SystemMain=" + form["SystemMain"] + ",Sytem Sub=" + form["SystemSub"]);
            //取得備標品項與詢價資料
            List<COMPARASION_DATA> lst = service.getComparisonData(form["id"], form["typeCode1"], form["typeCode2"], form["SystemMain"], form["SystemSub"]);
            log.Info("get Records=" + lst.Count);
            //產生畫面
            return PartialView(lst);
        }
        //更新單項成本資料
        public string UpdateCost4Item()
        {
            log.Info("ProjectItemID=" + Request["pitmid"] + ",Cost=" + Request["price"]);
            try
            {
                decimal cost = decimal.Parse(Request["price"]);
                service.updateCostFromQuote(Request["pitmid"], cost);
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
                return "更新失敗(請檢查資料格式是否有誤)!!";
            }
            return "更新成功!!";
        }
        //依據詢價單內容，更新標單所有單價
        public string BatchUpdateCost(string formid)
        {
            log.Info("formid=" + Request["formid"] );
            int i = service.batchUpdateCostFromQuote(Request["formid"]);         
            return "更新成功!!";
        }
    }
}