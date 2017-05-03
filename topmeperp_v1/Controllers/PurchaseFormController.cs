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
using System.Data;


namespace topmeperp.Controllers
{
    public class PurchaseFormController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));
        PurchaseFormService service = new PurchaseFormService();

        // GET: PurchaseForm
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "專案執行");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }

        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname, string status)
        {
            if (projectname != null)
            {
                log.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%' AND STATUS=@status;",
                         new SqlParameter("projectname", projectname), new SqlParameter("status", status)).ToList();
                }
                log.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }

        public ActionResult FormIndex(string id)
        {
            log.Info("purchase form index : projectid=" + id);
            ViewBag.projectId = id;
            return View();
        }
        // POST : Search
        [HttpPost]
        public ActionResult FormIndex(FormCollection f)
        {
            PlanService s = new PlanService();
            log.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            List<topmeperp.Models.PLAN_ITEM> lstProject = s.getPlanItem(Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            return View("FormIndex", lstProject);
        }
        //Create Purchasing Form
        [HttpPost]
        public ActionResult Create(PLAN_SUP_INQUIRY qf)
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
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            qf.PROJECT_ID = Request["prjId"];
            qf.CREATE_ID = u.USER_ID;
            qf.CREATE_DATE = DateTime.Now;
            qf.OWNER_NAME = uInfo.USER_NAME;
            qf.OWNER_EMAIL = uInfo.EMAIL;
            qf.OWNER_TEL = uInfo.TEL;
            qf.OWNER_FAX = uInfo.FAX;
            PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
            string fid = service.newPlanForm(qf, lstItemId);
            //產生採購詢價單實體檔案(先註解掉，因為空白詢價單不用產生實體檔，
            //樣本轉廠商採購單時再產生即可)
            //service.getInqueryForm(fid);
            //PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
            //poi.exportExcel4po(service.formInquiry, service.formInquiryItem);
            return Redirect("FormMainPage?id=" + qf.PROJECT_ID);
            //return RedirectToAction("InquiryMainPage","Inquiry", qf.PROJECT_ID);
        }
        public ActionResult FormMainPage(string id)
        {
            log.Info("purchase form by projectID=" + id);
            PurchaseFormModel formData = new PurchaseFormModel();
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                formData.planTemplateForm = service.getFormTemplateByProject(id);
                formData.planFormFromSupplier = service.getFormByProject(id);
            }
            return View(formData);
        }
        //顯示單一詢價單、報價單功能
        public ActionResult SinglePrjForm(string id)
        {
            log.Info("http get mehtod:" + id);
            PurchaseFormDetail singleForm = new PurchaseFormDetail();
            service.getInqueryForm(id);
            singleForm.planForm = service.formInquiry;
            singleForm.planFormItem = service.formInquiryItem;
            singleForm.prj = service.getProjectById(singleForm.planForm.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            //取得供應商資料
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            List<SelectListItem> selectSupplier = new List<SelectListItem>();
            foreach (string itm in service.getSupplier())
            {
                log.Debug("Supplier=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectSupplier.Add(selectI);
                }
            }
            // selectSupplier.Add(empty);
            ViewBag.Supplier = selectSupplier;
            return View(singleForm);
        }

        public String UpdateFormName(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得空白詢價單名稱
            string[] formId = form.Get("inquiryformid").Split(',');
            string[] formName = form.Get("formname").Split(',');
            List<PLAN_SUP_INQUIRY> lstItem = new List<PLAN_SUP_INQUIRY>();
            for (int j = 0; j < formId.Count(); j++)
            {
                PLAN_SUP_INQUIRY item = new PLAN_SUP_INQUIRY();
                item.INQUIRY_FORM_ID = formId[j];
                item.FORM_NAME = formName[j];
                log.Debug("plan form id =" + item.INQUIRY_FORM_ID + "，form name =" + item.FORM_NAME);
                lstItem.Add(item);
            }
            int i = service.addFormName(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新空白詢價單名稱成功";
            }
            return msg;
        }
    
        public String UpdatePrjForm(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            PLAN_SUP_INQUIRY fm = new PLAN_SUP_INQUIRY();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            fm.SUPPLIER_ID = form.Get("Supplier").Substring(7).Trim();
            fm.PROJECT_ID = form.Get("projectid").Trim();
            fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.CREATE_ID = loginUser.USER_ID;
            fm.CREATE_DATE = DateTime.Now;
            TND_SUPPLIER s = service.getSupplierInfo(form.Get("Supplier").Substring(0, 7).Trim());
            fm.CONTACT_NAME = s.CONTACT_NAME;
            fm.CONTACT_EMAIL = s.CONTACT_EMAIL;
            string[] lstItemId = form.Get("formitemid").Split(',');
            log.Info("select count:" + lstItemId.Count());
            var j = 0;
            for (j = 0; j < lstItemId.Count(); j++)
            {
                log.Info("item_list return No.:" + lstItemId[j]);
            }
            PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
            string i = service.addSupplierForm(fm, lstItemId);
            //產生廠商詢價單實體檔案
            service.getInqueryForm(i);
            PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
            poi.exportExcel4po(service.formInquiry, service.formInquiryItem);
            if (i == "")
            {
                msg = service.message;
            }
            else
            {
                msg = "新增供應商採購詢價單成功";
            }

            log.Info("Request:FORM_NAME=" + form["formname"] + "SUPPLIER_NAME =" + form["Supplier"]);
            return msg;
        }
        //更新採購廠商詢價單資料
        public String RefreshPrjForm(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            PLAN_SUP_INQUIRY fm = new PLAN_SUP_INQUIRY();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            fm.PROJECT_ID = form.Get("projectid").Trim();
            fm.SUPPLIER_ID = form.Get("supplier").Trim();
            fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.CONTACT_NAME = form.Get("inputcontact").Trim();
            fm.CONTACT_EMAIL = form.Get("inputemail").Trim();
            fm.INQUIRY_FORM_ID = form.Get("inputformnumber").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.CREATE_ID = form.Get("createid").Trim();
            fm.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            fm.MODIFY_ID = loginUser.USER_ID;
            fm.MODIFY_DATE = DateTime.Now;
            string formid = form.Get("inputformnumber").Trim();

            string[] lstItemId = form.Get("formitemid").Split(',');
            string[] lstPrice = form.Get("formunitprice").Split(',');
            List<PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
                item.INQUIRY_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstPrice[j].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                }
                log.Debug("Item No=" + item.INQUIRY_ITEM_ID + "=" + item.ITEM_UNIT_PRICE);
                lstItem.Add(item);
            }
            int i = service.refreshPlanSupplierForm(formid, fm, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新供應商採購詢價單成功，INQUIRY_FORM_ID =" + formid;
            }

            log.Info("Request: INQUIRY_FORM_ID = " + formid + "FORM_NAME =" + form["formname"] + "SUPPLIER_NAME =" + form["supplier"]);
            return msg;
        }

        //上傳廠商報價單
        public string FileUpload(HttpPostedFileBase file)
        {
            log.Info("Upload purchase form from supplier:" + Request["projectid"]);
            string projectid = Request["projectid"];
            //上傳至廠商報價單目錄
            if (null != file && file.ContentLength != 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid + "/" + ContextService.quotesFolder, fileName);
                file.SaveAs(path);
                log.Info("Parser Excel File Begin:" + file.FileName);
                PurchaseFormtoExcel quoteFormService = new PurchaseFormtoExcel();
                try
                {
                    quoteFormService.convertInquiry2Plan(path, projectid);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                }
                int i = 0;
                //如果詢價單編號為空白，新增詢價單資料，否則更新相關詢價單資料-new
                log.Debug("Parser Excel File Finish!");
                if (null != quoteFormService.form.INQUIRY_FORM_ID && quoteFormService.form.INQUIRY_FORM_ID != "")
                {
                    log.Info("Update Plan Form for Inquiry:" + quoteFormService.form.INQUIRY_FORM_ID);
                    i = service.refreshPlanSupplierForm(quoteFormService.form.INQUIRY_FORM_ID, quoteFormService.form, quoteFormService.formItems);
                }
                else
                {
                    log.Info("Create New Plan Form for Inquiry:");
                    i = service.createPlanFormFromSupplier(quoteFormService.form, quoteFormService.formItems);
                }
                log.Info("add plan supplier form record count=" + i);
            }
            return "檔案匯入成功!!";
        }
        //批次產生採購空白詢價單
        public string createPlanEmptyForm()
        {
            log.Info("project id=" + Request["projectid"]);
            SYS_USER u = (SYS_USER)Session["user"];
            int i = service.createPlanEmptyForm(Request["projectid"], u);
            return "共產生 " + i + "空白詢價單樣本!!";
        }
        public void downLoadInquiryForm()
        {
            string formid = Request["formid"];
            service.getInqueryForm(formid);
            if (null != service.formInquiry)
            {
                PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
                poi.exportExcel4po(service.formInquiry, service.formInquiryItem);
                Response.Clear();
                Response.Charset = "utf-8";
                Response.ContentType = "text/xls";
                Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", service.formInquiry.INQUIRY_FORM_ID + ".xlsx"));
                //"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx"
                Response.WriteFile(poi.outputPath + "\\" + service.formInquiry.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + service.formInquiry.INQUIRY_FORM_ID + ".xlsx");
                Response.End();
            }
        }
    }
}
