using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;
using System.Data;

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
            TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
            string fid = s.newForm(qf, lstItemId);
            //產生詢價單實體檔案-old
            service.getInqueryForm(fid);
            InquiryFormToExcel poi = new InquiryFormToExcel();
            poi.exportExcel(service.formInquiry, service.formInquiryItem);
            return Redirect("InquiryMainPage?id=" + qf.PROJECT_ID);
            //return RedirectToAction("InquiryMainPage","Inquiry", qf.PROJECT_ID);
        }
        //顯示單一詢價單、報價單功能
        public ActionResult SinglePrjForm(string id)
        {
            log.Info("http get mehtod:" + id);
            InquiryFormDetail singleForm = new InquiryFormDetail();
            service.getInqueryForm(id);
            singleForm.prjForm = service.formInquiry;
            singleForm.prjFormItem = service.formInquiryItem;
            singleForm.prj = service.getProjectById(singleForm.prjForm.PROJECT_ID);
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
        public String UpdatePrjForm(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            TND_PROJECT_FORM fm = new TND_PROJECT_FORM();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            fm.SUPPLIER_ID = form.Get("Supplier").Substring(7).Trim();
            fm.PROJECT_ID = form.Get("projectid").Trim();
            fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            TND_SUPPLIER s = service.getSupplierInfo(form.Get("Supplier").Substring(0, 7).Trim());
            fm.CONTACT_NAME = s.CONTACT_NAME;
            fm.CONTACT_EMAIL = s.CONTACT_EMAIL;
            fm.CREATE_ID = loginUser.USER_ID;
            fm.CREATE_DATE = DateTime.Now;
            string[] lstItemId = form.Get("formitemid").Split(',');
            log.Info("select count:" + lstItemId.Count());
            var j = 0;
            for (j = 0; j < lstItemId.Count(); j++)
            {
                log.Info("item_list return No.:" + lstItemId[j]);
            }
            TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
            string i = service.addNewSupplierForm(fm, lstItemId);
            //產生廠商詢價單實體檔案
            service.getInqueryForm(i);
            InquiryFormToExcel poi = new InquiryFormToExcel();
            poi.exportExcel(service.formInquiry, service.formInquiryItem);
            if (i == "")
            {
                msg = service.message;
            }
            else
            {
                msg = "新增供應商詢價單成功";
            }

            log.Info("Request:FORM_NAME=" + form["formname"] + "SUPPLIER_NAME =" + form["Supplier"]);
            return msg;
        }
        //更新廠商詢價單資料
        public String RefreshPrjForm(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            TND_PROJECT_FORM fm = new TND_PROJECT_FORM();
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
            fm.FORM_ID = form.Get("inputformnumber").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.CREATE_ID = form.Get("createid").Trim();
            fm.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            fm.MODIFY_ID = loginUser.USER_ID;
            fm.MODIFY_DATE = DateTime.Now;
            string formid = form.Get("inputformnumber").Trim();

            string[] lstItemId = form.Get("formitemid").Split(',');
            string[] lstPrice = form.Get("formunitprice").Split(',');
            List<TND_PROJECT_FORM_ITEM> lstItem = new List<TND_PROJECT_FORM_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
                item.FORM_ITEM_ID = int.Parse(lstItemId[j]);
                item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                log.Debug("Item No=" + item.FORM_ITEM_ID + "=" + item.ITEM_UNIT_PRICE);
                lstItem.Add(item);
            }

            int i = service.refreshSupplierForm(formid, fm, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新供應商詢價單成功，FORM_ID =" + formid;
            }

            log.Info("Request: FORM_ID = " + formid + "FORM_NAME =" + form["formname"] + "SUPPLIER_NAME =" + form["supplier"] );
            return msg;
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
            log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"]);
            //取得備標品項與詢價資料
            try
            {
                DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"]);
                @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";
                string htmlString = "<table class='table table-bordered'><tr>";
                //處理表頭
                for (int i = 1; i < 6; i++)
                {
                    log.Debug("column name=" + dt.Columns[i].ColumnName);
                    htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                }
                //處理供應商表頭
                Dictionary<string, COMPARASION_DATA> dirSupplierQuo = service.dirSupplierQuo;
                log.Debug("Column Count=" + dt.Columns.Count);
                for (int i = 6; i < dt.Columns.Count; i++)
                {
                    log.Debug("column name=" + dt.Columns[i].ColumnName);
                    string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                    //<a href="/Inquiry/SinglePrjForm/@item.FORM_ID" target="_blank">@item.FORM_ID</a>
                    decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;

                    htmlString = htmlString + "<th>" + tmpString[0] + "(" + tAmount + ")" +
                        "<button type='button' class='btn-xs' onclick=\"clickSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                        "<button type='button' class='btn-xs'><a href='/Inquiry/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a>" +
                        "</button>";
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
                    //單價欄位  <input type='text' id='cost_@item.PROJECT_ITEM_ID' name='cost_@item.PROJECT_ITEM_ID' size='5' />
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
                        //<td><button class="btn-link" onclick="clickPrice('@item.PROJECT_ITEM_ID', '@item.QUOTATION_PRICE')">@item.QUOTATION_PRICE</button> </td>
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
            log.Info("formid=" + Request["formid"]);
            int i = service.batchUpdateCostFromQuote(Request["formid"]);
            return "更新成功!!";
        }
        //成本分析
        public ActionResult costAnalysis(string id)
        {
            //產生成本分析Excel 並以固定檔案供使用者下載使用
            ViewBag.projectid = id;
            log.Info("Cost Analysis for projectid=" + id);
            CostAnalysisOutput excel = new CostAnalysisOutput();
            excel.exportExcel(id);
            ViewBag.url = "/UploadFile/" + id + "/" + id + "_CostAnalysis.xlsx";
            return View();
        }
    }
}
