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
using System.Web.Script.Serialization;
using Newtonsoft.Json;


namespace topmeperp.Controllers
{
    public class CashFlowController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        PurchaseFormService service = new PurchaseFormService();

        // GET: CashFlow 
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<CashFlowFunction> lstCashFlow = null;
            lstCashFlow = service.getCashFlow();
            return View(lstCashFlow);
        }

        //取得特定日期收入明細
        public ActionResult CashInFlowItem(string id)
        {
            List<PLAN_ACCOUNT> CashInFlow = null;
            CashInFlow = service.getDebitByDate(id);
            return View(CashInFlow);
        }

        //取得特定日期支出明細
        public ActionResult CashOutFlowItem(string id)
        {
            List<PLAN_ACCOUNT> CashOutFlow = null;
            CashOutFlow = service.getCreditByDate(id);
            return View(CashOutFlow);
        }

        public ActionResult ExpenseBudget()
        {
            logger.Info("Access to Expense Budget Page !!");
            List<ExpenseBudgetSummary> ExpBudget = null;
            ExpenseBudgetSummary Amt = null;
            if (null != Request["budgetyear"])
            {
                ExpBudget = service.getExpBudgetByYear(int.Parse(Request["budgetyear"]));
                Amt = service.getTotalExpBudgetAmount(int.Parse(Request["budgetyear"]));
                TempData["TotalAmt"] = Amt.TOTAL_BUDGET;
            }
            TempData["budgetYear"] = Request["budgetyear"];
            return View(ExpBudget);
        }

        public ActionResult Search()
        {
            List<ExpenseBudgetSummary> ExpBudget = null;
            ExpenseBudgetSummary Amt = null;
            if (null != Request["budgetyear"])
            {
                ExpBudget = service.getExpBudgetByYear(int.Parse(Request["budgetyear"]));
                Amt = service.getTotalExpBudgetAmount(int.Parse(Request["budgetyear"]));
                TempData["TotalAmt"] = String.Format("{0:#,##0.#}", Amt.TOTAL_BUDGET);
            }
            TempData["budgetYear"] = Request["budgetyear"];
            return View("ExpenseBudget", ExpBudget);
        }
        /// <summary>
        /// 下載公司費用預算填寫表
        /// </summary>
        public void downLoadExpBudgetForm()
        {
            ExpBudgetFormToExcel poi = new ExpBudgetFormToExcel();
            //檔案位置
            string fileLocation = poi.exportExcel();
            //檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted-Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
            string filename = HttpUtility.UrlEncode(Path.GetFileName(fileLocation));
            Response.Clear();
            Response.Charset = "utf-8";
            Response.ContentType = "text/xls";
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
            ///"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx"
            Response.WriteFile(fileLocation);
            Response.End();
        }
        //上傳公司費用預算
        [HttpPost]
        public ActionResult uploadExpBudgetTable(HttpPostedFileBase fileBudget)
        {
            int budgetYear = int.Parse(Request["year"]);
            logger.Info("Upload Expense Budget Table for budget year =" + budgetYear);
            string message = "";

            if (null != fileBudget && fileBudget.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileBudget.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileBudget.FileName);
                var path = Path.Combine(ContextService.strUploadPath, fileName);
                logger.Info("save excel file:" + path);
                fileBudget.SaveAs(path);
                //2.2 開啟Excel 檔案
                logger.Info("Parser Excel File Begin:" + fileBudget.FileName);
                ExpBudgetFormToExcel budgetservice = new ExpBudgetFormToExcel();
                budgetservice.InitializeWorkbook(path);
                //解析預算數量
                List<FIN_EXPENSE_BUDGET> lstExpBudget = budgetservice.ConvertDataForExpBudget(budgetYear);
                //2.3 記錄錯誤訊息
                message = budgetservice.errorMessage;
                //2.4
                logger.Info("Delete FIN_EXPENSE_BUDGET By Year");
                service.delExpBudgetByYear(budgetYear);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All FIN_EXPENSE_BUDGET to DB");
                service.refreshExpBudget(lstExpBudget);
                message = message + "<br/>資料匯入完成 !!";
            }
            TempData["result"] = message;
            return RedirectToAction("ExpenseBudget");
        }

        //更新公司費用預算
        public String UpdateExpBudget(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Expense Budget Year =" + form["year"]);
            logger.Info("Delete FIN_EXPENSE_BUDGET By BUDGET_YEAR");
            service.delExpBudgetByYear(int.Parse(form["year"]));
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt").Split(',');
            string[] lst2 = form.Get("febAmt").Split(',');
            string[] lst3 = form.Get("marAmt").Split(',');
            string[] lst4 = form.Get("aprAmt").Split(',');
            string[] lst5 = form.Get("mayAmt").Split(',');
            string[] lst6 = form.Get("junAmt").Split(',');
            string[] lst7 = form.Get("julAmt").Split(',');
            string[] lst8 = form.Get("augAmt").Split(',');
            string[] lst9 = form.Get("sepAmt").Split(',');
            string[] lst10 = form.Get("octAmt").Split(',');
            string[] lst11 = form.Get("novAmt").Split(',');
            string[] lst12 = form.Get("decAmt").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<FIN_EXPENSE_BUDGET> lst = new List<FIN_EXPENSE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<FIN_EXPENSE_BUDGET> lstItem = new List<FIN_EXPENSE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    FIN_EXPENSE_BUDGET item = new FIN_EXPENSE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["year"]);
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshExpBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新公司預算費用成功，預算年度為 " + form["year"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["year"]);
            return msg;
        }

        public ActionResult OperatingExpense()
        {
            logger.Info("Access to Operating Expense Page !!");
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View();
        }

        public ActionResult SearchSubject()
        {
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["subject"]);
            string[] lstItemId = Request["subject"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            List<FIN_SUBJECT> SubjectChecked = null;
            SubjectChecked = service.getSubjectByChkItem(lstItemId);
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View("OperatingExpense", SubjectChecked);
        }

        [HttpPost]
        public ActionResult AddExpense(FIN_EXPENSE_FORM ef)
        {
            string[] lstSubject = Request["subject"].Split(',');
            string[] lstAmount = Request["expense_amount"].Split(',');
            string[] lstRemark = Request["item_remark"].Split(',');
            //建立公司費用單號
            logger.Info("create new Operating Expense Form");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            ef.PAYMENT_DATE = Convert.ToDateTime(Request["paymentdate"]);
            ef.OCCURRED_YEAR = int.Parse(Request["occurreddate"].Substring(0, 4));
            ef.OCCURRED_MONTH = int.Parse(Request["occurreddate"].Substring(5, 2));
            ef.CREATE_DATE = DateTime.Now;
            ef.CREATE_ID = uInfo.USER_ID;
            ef.REMARK = Request["remark"];
            ef.STATUS = 10;
            string fid = service.newExpenseForm(ef);
            //建立公司費用單明細
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                item.ITEM_REMARK = lstRemark[j];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Operating Expense Subject =" + item.FIN_SUBJECT_ID + "， and Amount = " + item.AMOUNT);
                item.EXP_FORM_ID = fid;
                logger.Debug("Item EX form id =" + item.EXP_FORM_ID);
                lstItem.Add(item);
            }
            int i = service.AddExpenseItems(lstItem);
            logger.Debug("Item Count =" + i);
            return Redirect("SingleEXPForm?id=" + fid);
        }

        //顯示單一公司營業費用單/工地費用單功能
        public ActionResult SingleEXPForm(string id)
        {
            logger.Info("http get mehtod:" + id);
            OperatingExpenseModel singleForm = new OperatingExpenseModel();
            service.getEXPByExpId(id);
            singleForm.finEXP = service.formEXP;
            singleForm.finEXPItem = service.EXPItem;
            singleForm.planEXPItem = service.siteEXPItem;
            logger.Debug("Expense Year:" + singleForm.finEXP.OCCURRED_YEAR);
            return View(singleForm);
        }

        //更新費用單
        public String UpdateEXP(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得費用單資料
            FIN_EXPENSE_FORM ef = new FIN_EXPENSE_FORM();
            ef.OCCURRED_YEAR = int.Parse(form.Get("year").Trim());
            ef.OCCURRED_MONTH = int.Parse(form.Get("month").Trim());
            ef.REMARK = form.Get("remark").Trim();
            ef.CREATE_ID = form.Get("createid").Trim();
            ef.PAYMENT_DATE = Convert.ToDateTime(form.Get("paymentdate"));
            ef.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            ef.STATUS = int.Parse(form.Get("status").Trim());
            ef.MODIFY_DATE = DateTime.Now;
            ef.EXP_FORM_ID = form.Get("formnumber").Trim();
            ef.PROJECT_ID = form.Get("projectid").Trim(); 
            string[] lstSubject = form.Get("subject").Split(',');
            string[] lstRemark = form.Get("item_remark").Split(',');
            string[] lstAmount = form.Get("amount").Split(',');
            string formid = form.Get("formnumber").Trim();
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                if (lstRemark[j].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[j];
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Debug("Subject Id =" + item.FIN_SUBJECT_ID + ", Amount =" + item.AMOUNT);
                lstItem.Add(item);
            }
            int i = service.refreshEXPForm(formid, ef, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else if(form["projectid"] != null && form["projectid"] != "")
            {
                msg = "更新工地費用單成功";
            }
            else
            {
                msg = "更新公司營業費用單成功";
            }
            logger.Info("Request: 更新公司營業費用/工地費用單訊息 = " + msg);
            return msg;
        }

        public String UpdateEXPStatusById(FormCollection form)
        {
            //取得費用單編號
            logger.Info("form:" + form.Count);
            logger.Info("EXP form Id:" + form["formnumber"]);
            string msg = "";
            FIN_EXPENSE_FORM ef = new FIN_EXPENSE_FORM();
            ef.OCCURRED_YEAR = int.Parse(form.Get("year").Trim());
            ef.OCCURRED_MONTH = int.Parse(form.Get("month").Trim());
            ef.REMARK = form.Get("remark").Trim();
            ef.CREATE_ID = form.Get("createid").Trim();
            ef.PAYMENT_DATE = Convert.ToDateTime(form.Get("paymentdate"));
            ef.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            ef.STATUS = int.Parse(form.Get("status").Trim());
            ef.MODIFY_DATE = DateTime.Now;
            ef.EXP_FORM_ID = form.Get("formnumber").Trim();
            ef.PROJECT_ID = form.Get("projectid").Trim();
            string[] lstSubject = form.Get("subject").Split(',');
            string[] lstRemark = form.Get("item_remark").Split(',');
            string[] lstAmount = form.Get("amount").Split(',');
            string formid = form.Get("formnumber").Trim();
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                if (lstRemark[j].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[j];
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Debug("Subject Id =" + item.FIN_SUBJECT_ID + ", Amount =" + item.AMOUNT);
                lstItem.Add(item);
            }
            int i = service.refreshEXPForm(formid, ef, lstItem);
            //更新費用單狀態
            logger.Info("Update Expense Form Status");
            //費用單(已送審) STATUS = 20
            int k = service.RefreshEXPStatusById(formid);
            if (k == 0)
            {
                msg = service.message;
            }
            else if(form["projectid"] != null && form["projectid"] != "")
            {
                msg = "工地費用單已送審";
            }
            else
            {
                msg = "公司營業費用單已送審";
            }
            return msg;
        }

        //費用單查詢
        public ActionResult ExpenseForm(string id)
        {
            logger.Info("Search For Expense Form !!");
            //費用單草稿
            int status = 20;
            if (Request["status"] == null || Request["status"] == "")
            {
                status = 10;
            }
            if (id != null && id != "")
            {
                TND_PROJECT p = service.getProjectById(id);
                ViewBag.projectName = p.PROJECT_NAME;
                ViewBag.projectid = id;
            }
            else
            {
                id = "";
                ViewBag.projectid = "";
            }
            List<OperatingExpenseFunction> lstEXP = service.getEXPListByExpId(Request["occurred_date"], Request["subjectname"], Request["expid"], status, id);
            return View(lstEXP);
        }

        public ActionResult SearchEXP()
        {
            //logger.Info("occurred_date =" + Request["occurred_date"] + "subjectname =" + Request["subjectname"] + ", expid =" + Request["expid"] + ", status =" + int.Parse(Request["status"] + ", projectid =" + Request["id"]));
            List<OperatingExpenseFunction> lstEXP = service.getEXPListByExpId(Request["occurred_date"], Request["subjectname"], Request["expid"], int.Parse(Request["status"]), Request["id"]);
            ViewBag.SearchResult = "共取得" + lstEXP.Count + "筆資料";
            if (Request["id"] != null && Request["id"] != "")
            {
                TND_PROJECT p = service.getProjectById(Request["id"]);
                ViewBag.projectName = p.PROJECT_NAME;
                ViewBag.projectid = Request["id"];
            }
            return View("ExpenseForm", lstEXP);
        }

        public String RejectEXPById(FormCollection form)//須設定角色來鎖定每個button的權限(目前還未處理)
        {
            //取得費用單編號
            logger.Info("EXP form Id:" + form["formnumber"]);
            //更新費用單狀態
            logger.Info("Reject Expense Form ");
            string formid = form.Get("formnumber").Trim();
            //費用單(已退件) STATUS = 0
            string msg = "";
            int i = service.RejectEXPByExpId(formid);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "費用單已退回";
            }
            return msg;
        }

        public String PassEXPById(FormCollection form)
        {
            //取得費用單編號
            logger.Info("EXP form Id:" + form["formnumber"]);
            //更新費用單狀態
            logger.Info("Pass Expense Form ");
            string formid = form.Get("formnumber").Trim();
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            string passid = uInfo.USER_ID;
            //費用單(主管已通過) STATUS = 30
            string msg = "";
            int i = service.PassEXPByExpId(formid, passid);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "費用單已核可";
            }
            return msg;
        }

        public String JournalById(FormCollection form)
        {
            //取得費用單編號
            logger.Info("EXP form Id:" + form["formnumber"]);
            //更新費用單狀態
            logger.Info("Journal For Operating Expense Form ");
            string formid = form.Get("formnumber").Trim();
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            string journalid = uInfo.USER_ID;
            //費用(已立帳) STATUS = 40
            string msg = "";
            int i = service.JournalByExpId(formid, journalid);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "費用單已核可";
            }
            return msg;
        }

        public String ApproveEXPById(FormCollection form)
        {
            //取得費用單編號
            logger.Info("EXP form Id:" + form["formnumber"]);
            //更新費用單狀態
            logger.Info("Approve Operating Expense Form ");
            string formid = form.Get("formnumber").Trim();
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            string approveid = uInfo.USER_ID;
            //費用單(已核可) STATUS = 50
            string msg = "";
            int i = service.ApproveEXPByExpId(formid, approveid);
            string k = service.AddAccountByExpId(formid, approveid);
            logger.Info("Add the Operating Expense Account To Plan Account Record, It's Form Id = " + k);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "費用單已核可";
            }
            return msg;
        }

        public String UpdateAccountStatus(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            int i = 0;
            string[] lstForm = form.Get("formid").Split(',');
            List<PLAN_ACCOUNT> lstItem = new List<PLAN_ACCOUNT>();
            if (form.Get("status") != null)
            {
                string[] lstStatus = form.Get("status").Split(',');
                for (int j = 0; j < lstForm.Count(); j++)
                {
                    PLAN_ACCOUNT item = new PLAN_ACCOUNT();
                    item.ACCOUNT_FORM_ID = lstForm[j];
                    item.MODIFY_DATE = DateTime.Now;
                    if (lstStatus[j].ToString() == "")
                    {
                        item.STATUS = 10;
                    }
                    else
                    {
                        item.STATUS = 0;
                    }
                    logger.Debug("Acount Form Id =" + item.ACCOUNT_FORM_ID + ", Status =" + item.STATUS);
                    lstItem.Add(item);
                }
            }
            else
            {
                for (int j = 0; j < lstForm.Count(); j++)
                {
                    PLAN_ACCOUNT item = new PLAN_ACCOUNT();
                    item.ACCOUNT_FORM_ID = lstForm[j];
                    item.STATUS = 10;
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Debug("Acount Form Id =" + item.ACCOUNT_FORM_ID + ", Status =" + item.STATUS);
                    lstItem.Add(item);
                }
            }
            i = service.refreshAccountStatus(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "帳款支付狀態已更新";
            }
            return msg;
        }

        //會計立帳進入畫面
        public ActionResult FormForJournal()
        {
            logger.Info("Access to Form For Journal !!");
            //公司需立帳之帳款(即會計審核)
            int status = 30;
            ViewBag.forJournal = "會計立帳"; 
            List<OperatingExpenseFunction> lstEXP = service.getEXPListByExpId(Request["occurred_date"], Request["subjectname"], Request["expid"], status, Request["id"]);
            return View(lstEXP);
        }

        public ActionResult SearchForm4Journal()
        {
            logger.Info("occurred_date =" + Request["occurred_date"] + ", subjectname =" + Request["subjectname"] + ", expid =" + Request["expid"] + ", status =" + int.Parse(Request["status"]) + ", projectid =" + Request["id"]);
            List<OperatingExpenseFunction> lstEXP = service.getEXPListByExpId(Request["occurred_date"], Request["subjectname"], Request["expid"], int.Parse(Request["status"]), Request["id"]);
            //ViewBag.SearchResult = "共取得" + lstEXP.Count + "筆資料";
            return View("FormForJournal", lstEXP);
        }

        //修改帳款支付日期
        public ActionResult PlanAccount()
        {
            logger.Info("Search For Account To Update Its Payment Date !!");
            return View();
        }

        public ActionResult ShowPlanAccount()
        {
            logger.Info("payment_date =" + Request["payment_date"] + ", projectname =" + Request["projectname"] + ", payee =" + Request["payee"] + ", account_type =" + Request["account_type"]);
            List<PlanAccountFunction> lstAccount = service.getPlanAccount(Request["payment_date"], Request["projectname"], Request["payee"], Request["account_type"]);
            ViewBag.SearchResult = "共取得" + lstAccount.Count + "筆資料";
            return PartialView(lstAccount);
        }

        public string getPlanAccountItem(string itemid)
        {
            logger.Info("get plan account item by id=" + itemid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPlanAccountItem(itemid));
            logger.Info("plan account item  info=" + itemJson);
            return itemJson;
        }

        public String updatePlanAccountItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_ACCOUNT item = new PLAN_ACCOUNT();
            item.PROJECT_ID = form["project_id"];
            item.PLAN_ACCOUNT_ID = int.Parse(form["plan_account_id"]);
            item.CONTRACT_ID = form["contract_id"];
            item.ACCOUNT_FORM_ID = form["account_form_id"];
            item.PAYMENT_DATE = Convert.ToDateTime(form.Get("date"));
            try
            {
                item.AMOUNT = decimal.Parse(form["amount"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ACCOUNT_ID + " not amount:" + ex.Message);
            }
            item.ACCOUNT_TYPE = form["type"];
            logger.Debug("account type = " + form["type"]);
            item.ISDEBIT = form["isdebit"];
            item.STATUS = int.Parse(form["unRecordedFlag"]);
            item.CREATE_ID = form["create_id"];
            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.MODIFY_ID = loginUser.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            int i = 0;
            i = service.updatePlanAccountItem(item);
            if (i == 0) { msg = service.message; }
            return msg;
        }
    }
}
