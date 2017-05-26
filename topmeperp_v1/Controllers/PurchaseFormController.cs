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

            log.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            List<topmeperp.Models.PLAN_ITEM> lstProject = service.getPlanItem(Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"]);
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
            log.Info("purchase form by projectID =" + id + ",status=" + Request["status"]);
            PurchaseFormModel formData = new PurchaseFormModel();
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                formData.planTemplateForm = service.getFormTemplateByProject(id);
                formData.planFormFromSupplier = service.getFormByProject(id, Request["status"]); 
            }
            ViewBag.Status = "有效";
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
            if (form.Get("inquiryformid") != null)
            {
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
            else
            {
                return "無詢價單名稱需要輸入，所以無法更新詢價單名稱";
            }
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
            if (form.Get("status").Trim() == "")
            {
                fm.STATUS = null;
            }
            else
            {
                fm.STATUS = form.Get("status").Trim();
            }
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
                log.Debug("Item No=" + item.INQUIRY_ITEM_ID + ", Price =" + item.ITEM_UNIT_PRICE );
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
                //"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.INQUIRY_FORM_ID + ".xlsx"
                Response.WriteFile(poi.outputPath + "\\" + service.formInquiry.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + service.formInquiry.INQUIRY_FORM_ID + ".xlsx");
                Response.End();
            }
        }
        //議約採購功能主頁
        public ActionResult PurchaseMain(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);
            //取得專案基本資料
            ViewBag.id = id;
            PurchaseFormService service = new PurchaseFormService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得未決標需議價之詢價單資料
            List<purchasesummary> lstforms = service.getPurchaseForm4Offer(id, Request["formname"]);
            BudgetDataService bs = new BudgetDataService();
            DirectCost iteminfo = bs.getItemBudget(id, Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formname"]);
            ViewBag.itembudget = iteminfo.ITEM_BUDGET;
            ViewBag.SearchResult = "共取得" + lstforms.Count + "筆資料";
            return View(lstforms);
        }
        public ActionResult Search()
        {
            List<purchasesummary> lstforms = service.getPurchaseForm4Offer(Request["id"], Request["formname"]);
            BudgetDataService bs = new BudgetDataService();
            DirectCost iteminfo = bs.getItemBudget(Request["id"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formname"]);
            ViewBag.itembudget = iteminfo.ITEM_BUDGET;
            ViewBag.SearchResult = "共取得" + lstforms.Count + "筆資料";
            return View("PurchaseMain", lstforms);
        }

        //採購比價功能資料頁
        public ActionResult Comparison(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);

            //取得專案基本資料
            PurchaseFormService service = new PurchaseFormService();
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
            log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"] + ",Form Name=" + Request["formName"]);
            //取得備標品項與詢價資料
            try
            {
                DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"]);
                @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";
                string htmlString = "<table class='table table-bordered'><tr>";
                //處理表頭
                for (int i = 1; i < 6; i++)
                {
                    log.Debug("column name=" + dt.Columns[i].ColumnName);
                    htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                }
                //處理供應商表頭
                Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = service.dirSupplierQuo;
                log.Debug("Column Count=" + dt.Columns.Count);
                List<string> list = new List<string>();
                for (int i = 6; i < dt.Columns.Count; i++)
                {
                    log.Debug("column name=" + dt.Columns[i].ColumnName);
                    string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                    //<a href="/PurchaseForm/SinglePrjForm/@item.INQUIRY_FORM_ID" target="_blank">@item.INQUIRY_FORM_ID</a>
                    decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                    string strAmout = string.Format("{0:C0}", tAmount);

                    htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                        "</td><button type='button' class='btn-xs' onclick=\"clickSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                        "</button>" + "<button type = 'button' class='btn-xs' onclick=\"removeSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></button>" +
                        "<button type='button' class='btn-xs'><a href='/PurchaseForm/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a>" +
                        "<br/><tr><td style = 'text-align:center;background-color:yellow;' > " + strAmout + "</td>" +
                        "</tr></table></th>";
                }
                htmlString = htmlString + "</tr>";
                //處理資料表
                foreach (DataRow dr in dt.Rows)
                {
                    htmlString = htmlString + "<tr>";
                    for (int i = 1; i < 5; i++)
                    {
                        htmlString = htmlString + "<td>" + dr[i] + "</td>";
                    }
                    //單價欄位  <input type='text' id='cost_@item.INQUIRY_ITEM_ID' name='cost_@item.INQUIRY_ITEM_ID' size='5' />
                    //decimal price = decimal.Parse(dr[5].ToString());
                    if (dr[5].ToString() != "")
                    {
                        log.Debug("data row col 5=" + (decimal)dr[5]);
                        htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N0}", (decimal)dr[5]) + "' /></td>";
                    }
                    else
                    {
                        htmlString = htmlString + "<td></td>";
                    }
                    //String.Format("{0:C}", 0);
                    //處理報價資料
                    for (int i = 6; i < dt.Columns.Count; i++)
                    {
                        //<td><button class="btn-link" onclick="clickPrice('@item.INQUIRY_ITEM_ID', '@item.QUOTATION_PRICE')">@item.QUOTATION_PRICE</button> </td>
                        if (dr[i].ToString() != "")
                        {
                            htmlString = htmlString + "<td><button class='btn-link' onclick=\"clickPrice('" + dr[1] + "', '" + dr[i] + "')\">" + String.Format("{0:N0}", (decimal)dr[i]) + "</button> </td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                    }
                    htmlString = htmlString + "</tr>";
                }
                htmlString = htmlString + "</table>";
                //加入預算資料
                BudgetDataService bs = new BudgetDataService();
                DirectCost iteminfo = bs.getItemBudget(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"]); 
                ViewBag.itembudget = iteminfo.ITEM_BUDGET;
                //產生畫面
                IHtmlString str = new HtmlString(htmlString);
                ViewBag.htmlString = str;
            }
            catch (Exception e)
            {
                log.Error("Ex" + e.Message);
                ViewBag.htmlString = e.Message;
            }
            return PartialView();
        }

        //移除議價資料
        public string RemoveSupplierForm(string formid)
        {
            log.Info("formid=" + Request["formid"]);
            int i = service.removeSuplplierFormFromQuote(Request["formid"]);
            return "更新成功!!";
        }
        //更新單項成本資料
        public string UpdateCost4Item()
        {
            log.Info("PLanItemID=" + Request["pitmid"] + ",Cost=" + Request["price"]);
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
        //依據詢價單內容，更新得標標單品項所有單價
        public string BatchUpdateCost(string formid)
        {
            log.Info("formid=" + Request["formid"]);
            int i = service.batchUpdateCostFromQuote(Request["formid"]);
            return "更新成功!!";
        }
        //取得採購合約資料
        public ActionResult PurchasingContract(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);
            //取得專案基本資料
            ViewBag.id = id;
            PurchaseFormService service = new PurchaseFormService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得採購合約財務資料
            List<plansummary> lstplanContract = service.getPlanContract(id);
            plansummary contractAmount = service.getPlanContractAmount(id);
            ViewBag.budget = contractAmount.TOTAL_BUDGET;
            ViewBag.cost = contractAmount.TOTAL_COST;
            ViewBag.revenue = contractAmount.TOTAL_REVENUE;
            ViewBag.profit = contractAmount.TOTAL_PROFIT;
            ViewBag.SearchResult = "共取得" + lstplanContract.Count + "筆資料";
            return View(lstplanContract);
        }

        //取得合約付款條件
        public string getPaymentTerms(string contractid)
        {
            PurchaseFormService service = new PurchaseFormService();
            log.Info("access the terms of payment by:" + Request["contractid"]);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPaymentTerms(contractid));
            log.Info("plan payment terms info=" + itemJson);
            return itemJson;
        }
        List<PLAN_ITEM> planitems = null;
        //取得採購遺漏項目
        public ActionResult PendingItems(string id)
        {
            log.Info("start project id=" + id);
            PurchaseFormService service = new PurchaseFormService();
            List<PLAN_ITEM> lstItem = new List<PLAN_ITEM>();
            planitems = service.getPendingItems(id);
            return View(planitems);
        }
    }
}
