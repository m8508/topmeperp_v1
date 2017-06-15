﻿using log4net;
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
    public class PlanController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        PlanService service = new PlanService();


        // GET: Plan
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "專案執行");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            //畫面上權限管理控制
            //頁面上使用ViewBag 定義開關\@ViewBag.F10005
            //由Session 取得權限清單
            List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
            //開關預設關閉
            @ViewBag.F10005 = "disabled";
            //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10005 = "";
            foreach (SYS_FUNCTION f in lstFunctions)
            {
                if (f.FUNCTION_ID == "F10005")
                {
                    @ViewBag.F10005 = "";
                }
            }
            return View(lstProject);
        }

        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname, string status)
        {
            if (projectname != null)
            {
                logger.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%' AND STATUS=@status;",
                         new SqlParameter("projectname", projectname), new SqlParameter("status", status)).ToList();
                }
                logger.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }

        //上傳得標後標單內容(用於標單內容有異動時)
        [HttpPost]
        public ActionResult uploadPlanItem(TND_PROJECT prj, HttpPostedFileBase file)
        {
            //1.取得專案編號
            string projectid = Request["projectid"];
            logger.Info("Upload plan items for projectid=" + projectid);
            string message = "";
            TND_PROJECT p = service.getProject(projectid);
            if (null != file && file.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + file.FileName);
                //2.1 將上傳檔案存檔
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                file.SaveAs(path);
                //2.2 解析Excel 檔案
                PlanItemFromExcel poiservice = new PlanItemFromExcel();
                poiservice.InitializeWorkbook(path);
                poiservice.ConvertDataForPlan(projectid);
                //2.3 記錄錯誤訊息
                message = message + "得標標單品項:共" + poiservice.lstPlanItem.Count + "筆資料，";
                message = message + "<a target=\"_blank\" href=\"/Plan/ManagePlanItem?id=" + projectid + "\"> 標單明細檢視畫面單</a><br/>" + poiservice.errorMessage;
                //        < button type = "button" class="btn btn-primary" onclick="location.href='@Url.Action("ManagePlanItem","Plan", new { id = @Model.tndProject.PROJECT_ID})'; ">標單明細</button>
                //2.4
                logger.Info("Delete PLAN_ITEM By Project ID");
                service.delAllItem();
                //2.5
                logger.Info("Add All PLAN_ITEM to DB");
                service.refreshItem(poiservice.lstPlanItem);
            }
            TempData["result"] = message;
            //PlanService ps = new PlanService();
            //var priId = ps.getBudgetById(projectid);
            //ViewBag.budgetdata = priId;
            //if (null != priId) { int k = service.updateBudgetToPlanItem(projectid); }
            return Redirect("ManagePlanItem?id=" + projectid);
        }
        /// <summary>
        /// 設定標單品項查詢條件
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public ActionResult ManagePlanItem(string id)
        {
            //傳入專案編號，
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("start project id=" + id);

            //取得專案基本資料fc
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
                logger.Debug("Main System=" + itm);
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
                logger.Debug("Sub System=" + itm);
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
        /// <summary>
        /// 取得標單明細資料
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public ActionResult ShowPlanItems(FormCollection form)
        {
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"]);
            List<PLAN_ITEM> lstItems = service.getPlanItem(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"], Request["supplier"]);
            ViewBag.Result = "共幾" + lstItems.Count + "筆資料";
            return PartialView(lstItems);
        }
        public string getPlanItem(string itemid)
        {
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("get plan item by id=" + itemid);
            //PLAN_ITEM  item = service.getPlanItem.getUser(userid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPlanItem(itemid));
            logger.Info("plan item  info=" + itemJson);
            return itemJson;
        }

        public String addPlanItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_ITEM item = new PLAN_ITEM();
            item.PROJECT_ID = form["project_id"];
            item.PLAN_ITEM_ID = form["plan_item_id"];
            item.ITEM_ID = form["item_id"];
            item.ITEM_DESC = form["item_desc"];
            item.ITEM_UNIT = form["item_unit"];
            try
            {
                item.ITEM_QUANTITY = decimal.Parse(form["item_quantity"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not quattity:" + ex.Message);
            }
            try
            {
                item.ITEM_FORM_QUANTITY = decimal.Parse(form["item_form_quantity"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not form quattity:" + ex.Message);
            }
            try
            {
                item.ITEM_UNIT_PRICE = decimal.Parse(form["item_unit_price"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not unit price:" + ex.Message);
            }

            item.TYPE_CODE_1 = form["type_code_1"];
            item.TYPE_CODE_2 = form["type_code_2"];

            item.SYSTEM_MAIN = form["system_main"];
            item.SYSTEM_SUB = form["system_sub"];
            try
            {
                item.EXCEL_ROW_ID = long.Parse(form["excel_row_id"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ITEM_ID + " not exce row id:" + ex.Message);
            }

            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.MODIFY_USER_ID = loginUser.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            PurchaseFormService service = new PurchaseFormService();
            int i = service.updatePlanItem(item);
            if (i == 0) { msg = service.message; }
            return msg;
        }

        //取得業主合約金額
        public ActionResult ContractForOwner(string id)
        {
            //傳入專案編號，
            logger.Info("start project id=" + id);
            //取得專案基本資料
            ViewBag.id = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得業主合約金額
            PlanRevenue contractAmount = service.getPlanRevenueById(id);
            ViewBag.Amount = contractAmount.PLAN_REVENUE;
            ViewBag.contractid = contractAmount.CONTRACT_ID;
            int i = service.addContractId4Owner(id);
            return View();
        }

        public ActionResult Budget(string id)
        {
            logger.Info("budget info for projectid=" + id);
            ViewBag.projectid = id;
            TnderProject service = new TnderProject();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得直接成本資料
            PlanService ps = new PlanService();
            var priId = ps.getBudgetById(id);
            ViewBag.budgetdata = priId;
            if (null == priId)
            //取得九宮格組合之直接成本資料
            {
                CostAnalysisDataService s = new CostAnalysisDataService();
                List<DirectCost> budget1 = s.getDirectCost(id); // 目前工資是使用工率還未轉成工資
                ViewBag.result = "共有" + (budget1.Count - 1) + "筆資料";
                return View(budget1);
            }
            //取得已寫入之九宮格組合預算資料
            BudgetDataService bs = new BudgetDataService();
            List<DirectCost> budget2 = bs.getBudget(id);
            DirectCost totalinfo = bs.getTotalCost(id);
            ViewBag.budget = totalinfo.TOTAL_BUDGET;
            ViewBag.cost = totalinfo.TOTAL_COST;
            ViewBag.p_cost = totalinfo.TOTAL_P_COST;
            ViewBag.result = "共有" + budget2.Count + "筆資料";
            return View(budget2);
        }
        /// <summary>
        /// 下載預算寫表
        /// </summary>
        public void downLoadBudgetForm()
        {
            string projectid = Request["projectid"];
            service.getProjectId(projectid);
            if (null != service.budgetTable)
            {
                BudgetFormToExcel poi = new BudgetFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(service.budgetTable, service.budgetTableItem);
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
        }
        //上傳預算
        [HttpPost]
        public ActionResult uploadBudgetTable(HttpPostedFileBase fileBudget)
        {
            string projectid = Request["projectid"];
            logger.Info("Upload Budget Table for projectid=" + projectid);
            string message = "";
            //檢查工率乘數是否存在
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(Request["projectid"]);
            ViewBag.projectWage = p.WAGE_MULTIPLIER;
            //if (null == ViewBag.projectWage) { throw new Exception(Request["projectid"] + "'s Wage Multiplier is not exist !!");}
            //檔案變數名稱(fileBudget)需要與前端畫面對應(view 的 file name and file id)
            if (null != fileBudget && fileBudget.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileBudget.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileBudget.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                fileBudget.SaveAs(path);
                //2.2 開啟Excel 檔案
                logger.Info("Parser Excel File Begin:" + fileBudget.FileName);
                BudgetFormToExcel budgetservice = new BudgetFormToExcel();
                budgetservice.InitializeWorkbook(path);
                //解析預算數量
                List<PLAN_BUDGET> lstBudget = budgetservice.ConvertDataForBudget(projectid);
                //2.3 記錄錯誤訊息
                message = budgetservice.errorMessage;
                //2.4
                logger.Info("Delete PLAN_BUDGET By Project ID");
                service.delBudgetByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All PLAN_BUDGET to DB");
                service.refreshBudget(lstBudget);
                message = message + "<br/>資料匯入完成 !!";
            }
            TempData["result"] = message;
            // 將預算寫入得標標單
            int k = service.updateBudgetToPlanItem(projectid);
            return RedirectToAction("Budget/" + projectid);
        }


        //寫入預算
        public String UpdateBudget(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lsttypecode = form.Get("code1").Split(',');
            string[] lsttypesub = form.Get("code2").Split(',');
            string[] lstCost = form.Get("inputtndratio").Split(',');
            string[] lstPrice = form.Get("inputbudget").Split(',');
            List<PLAN_BUDGET> lstItem = new List<PLAN_BUDGET>();
            for (int j = 0; j < lstPrice.Count(); j++)
            {
                PLAN_BUDGET item = new PLAN_BUDGET();
                item.PROJECT_ID = form["id"];
                if (lstPrice[j].ToString() == "")
                {
                    item.BUDGET_RATIO = null;
                }
                else
                {
                    item.BUDGET_RATIO = decimal.Parse(lstPrice[j]);
                }
                if (lstCost[j].ToString() == "")
                {
                    item.TND_RATIO = null;
                }
                else
                {
                    item.TND_RATIO = decimal.Parse(lstCost[j]);
                }
                logger.Info("Budget ratio =" + item.BUDGET_RATIO + ",Cost ratio =" + item.TND_RATIO);
                item.TYPE_CODE_1 = lsttypecode[j];
                item.TYPE_CODE_2 = lsttypesub[j];
                item.CREATE_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且九宮格組合為" + item.TYPE_CODE_1 + item.TYPE_CODE_2);
                lstItem.Add(item);
            }
            int i = service.addBudget(lstItem);
            // 將預算寫入得標標單
            int k = service.updateBudgetToPlanItem(form["id"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增預算資料成功，PROJECT_ID =" + form["id"];
            }

            logger.Info("Request:PROJECT_ID =" + form["id"]);
            return msg;
        }
        //修改預算
        public String RefreshBudget(string id, FormCollection form)
        {
            logger.Info("form:" + form.Count);
            id = Request["id"];
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lsttypecode = form.Get("code1").Split(',');
            string[] lsttypesub = form.Get("code2").Split(',');
            string[] lstPrice = form.Get("inputbudget").Split(',');
            List<PLAN_BUDGET> lstItem = new List<PLAN_BUDGET>();
            for (int j = 0; j < lstPrice.Count(); j++)
            {
                PLAN_BUDGET item = new PLAN_BUDGET();
                item.PROJECT_ID = form["id"];
                if (lstPrice[j].ToString() == "")
                {
                    item.BUDGET_RATIO = null;
                }
                else
                {
                    item.BUDGET_RATIO = decimal.Parse(lstPrice[j]);
                }
                logger.Info("Budget ratio =" + item.BUDGET_RATIO);
                item.TYPE_CODE_1 = lsttypecode[j];
                item.TYPE_CODE_2 = lsttypesub[j];
                item.MODIFY_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且九宮格組合為" + item.TYPE_CODE_1 + item.TYPE_CODE_2);
                lstItem.Add(item);
            }
            int i = service.updateBudget(id, lstItem);
            int k = service.updateBudgetToPlanItem(form["id"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新預算資料成功，PROJECT_ID =" + form["id"];
            }

            logger.Info("Request:PROJECT_ID =" + form["id"]);
            return msg;
        }
        //複合詢價單預算
        public ActionResult BudgetForForm(string id)
        {
            logger.Info("budget for form : projectid=" + id);
            ViewBag.projectId = id;
            return View();
        }
        [HttpPost]
        public ActionResult BudgetForForm(FormCollection f)
        {

            logger.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"] + ",formName=" + Request["formName"]);
            PurchaseFormService service = new PurchaseFormService();
            List<topmeperp.Models.PLAN_ITEM> lstProject = service.getPlanItem(Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"], Request["supplier"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            ViewBag.textCode1 = Request["textCode1"];
            ViewBag.textCode2 = Request["textCode2"];
            ViewBag.textSystemMain = Request["textSystemMain"];
            ViewBag.textSystemSub = Request["textSystemSub"];
            BudgetDataService bs = new BudgetDataService();
            DirectCost totalinfo = bs.getTotalCost(Request["projectid"]);
            DirectCost iteminfo = bs.getItemBudget(Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"]);
            ViewBag.budget = totalinfo.TOTAL_BUDGET;
            ViewBag.cost = totalinfo.TOTAL_COST;
            ViewBag.itembudget = iteminfo.ITEM_BUDGET;
            ViewBag.itemcost = iteminfo.ITEM_COST;
            return View("BudgetForForm", lstProject);
        }
        //修改Plan_Item 個別預算
        public String UpdatePlanBudget(string id, FormCollection form)
        {
            logger.Info("form:" + form.Count);
            id = Request["projectId"];
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lstplanitemid = form.Get("planitemid").Split(',');
            string[] lstBudget = form.Get("budgetratio").Split(',');
            List<PLAN_ITEM> lstItem = new List<PLAN_ITEM>();
            for (int j = 0; j < lstBudget.Count(); j++)
            {
                PLAN_ITEM item = new PLAN_ITEM();
                item.PROJECT_ID = form["id"];
                if (lstBudget[j].ToString() == "")
                {
                    item.BUDGET_RATIO = null;
                }
                else
                {
                    item.BUDGET_RATIO = decimal.Parse(lstBudget[j]);
                }
                logger.Info("Budget ratio =" + item.BUDGET_RATIO);
                item.PLAN_ITEM_ID = lstplanitemid[j];
                item.MODIFY_USER_ID = u.USER_ID;
                logger.Debug("Item Project id =" + item.PROJECT_ID + "且plan item id 為" + item.PLAN_ITEM_ID + "其項目預算為" + item.BUDGET_RATIO);
                lstItem.Add(item);
            }
            int i = service.updateItemBudget(id, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新預算資料成功，PROJECT_ID =" + Request["projectId"];
            }

            logger.Info("Request:PROJECT_ID =" + Request["projectId"]);
            return msg;
        }
    }
}
