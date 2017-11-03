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
using System.Data;
using System.Globalization;
using System.Reflection;

namespace topmeperp.Controllers
{
    public class MaterialManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));
        PurchaseFormService service = new PurchaseFormService();
        ProjectPlanService planService = new ProjectPlanService();

        // GET: MaterialManage
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

        //Plan Task 任務連結物料
        public ActionResult PlanTask(string id)
        {
            log.Debug("show sreen for apply for material");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.TreeString = planService.getProjectTask4Tree(id);
            return View();
        }

        //物料申購
        public ActionResult Application(FormCollection form)
        {
            log.Info("Access to Application page!!");
            ViewBag.projectid = form["projectid"];
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(form["projectid"]);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.applyDate = DateTime.Now;
            string[] deviceItemId = null;
            string[] fpItemId = null;
            string[] pepItemId = null;
            string[] lcpItemId = null;
            string[] pluItemId = null;
            string[] fwItemId = null;
            List<string> AllItemId = new List<string>();
            //取得使用者勾選任務ID
            if (null != form["map_device"])
            {
                log.Info("device task_list:" + Request["map_device"]);
                deviceItemId = Request["map_device"].ToString().Split(',');

                log.Info("select count:" + deviceItemId.Count());
                var i = 0;
                for (i = 0; i < deviceItemId.Count(); i++)
                {
                    log.Info("device task_list return No.:" + deviceItemId[i]);
                    AllItemId.Add(deviceItemId[i]);
                    //ViewBag.uid = lstItemId[i];
                }
            }
            if (null != form["map_fp"])
            {
                log.Info("fp task_list:" + Request["map_fp"]);
                fpItemId = Request["map_fp"].ToString().Split(',');

                log.Info("select count:" + fpItemId.Count());
                var i = 0;
                for (i = 0; i < fpItemId.Count(); i++)
                {
                    log.Info("fp task_list return No.:" + fpItemId[i]);
                    AllItemId.Add(fpItemId[i]);
                }
            }
            if (null != form["map_pep"])
            {
                log.Info("pep task_list:" + Request["map_pep"]);
                pepItemId = Request["map_pep"].ToString().Split(',');

                log.Info("select count:" + pepItemId.Count());
                var i = 0;
                for (i = 0; i < pepItemId.Count(); i++)
                {
                    log.Info("pep task_list return No.:" + pepItemId[i]);
                    AllItemId.Add(pepItemId[i]);
                }
            }
            if (null != form["map_lcp"])
            {
                log.Info("pep task_list:" + Request["map_lcp"]);
                lcpItemId = Request["map_lcp"].ToString().Split(',');

                log.Info("select count:" + lcpItemId.Count());
                var i = 0;
                for (i = 0; i < lcpItemId.Count(); i++)
                {
                    log.Info("lcp task_list return No.:" + lcpItemId[i]);
                    AllItemId.Add(lcpItemId[i]);
                }
            }
            if (null != form["map_plu"])
            {
                log.Info("plu task_list:" + Request["map_plu"]);
                pluItemId = Request["map_plu"].ToString().Split(',');

                log.Info("select count:" + pluItemId.Count());
                var i = 0;
                for (i = 0; i < pluItemId.Count(); i++)
                {
                    log.Info("plu task_list return No.:" + pluItemId[i]);
                    AllItemId.Add(pluItemId[i]);
                }
            }
            if (null != form["map_fw"])
            {
                log.Info("fw task_list:" + Request["map_fw"]);
                fwItemId = Request["map_fw"].ToString().Split(',');

                log.Info("select count:" + fwItemId.Count());
                var i = 0;
                for (i = 0; i < fwItemId.Count(); i++)
                {
                    log.Info("fw task_list return No.:" + fwItemId[i]);
                    AllItemId.Add(fwItemId[i]);
                }
            }
            if (null == form["map_device"] && null == form["map_fp"] && null == form["map_fw"] && null == form["map_pep"] && null == form["map_lcp"] && null == form["map_plu"])
            {
                TempData["result"] = "沒有選取要申購的項目名稱，請重新查詢後並勾選物料項目!";
                return Redirect("PlanTask?id=" + form["projectid"]);
            }
            else
            {
                List<PurchaseRequisition> lstPR = service.getPurchaseItemByMap(form["projectid"], AllItemId);
                return View(lstPR);
            }
        }

        //DOM 申購作業功能紐對應不同Action
        public class MultiButtonAttribute : ActionNameSelectorAttribute
        {
            public string Name { get; set; }
            public MultiButtonAttribute(string name)
            {
                this.Name = name;
            }
            public override bool IsValidName(ControllerContext controllerContext,
                string actionName, System.Reflection.MethodInfo methodInfo)
            {
                if (string.IsNullOrEmpty(this.Name))
                {
                    return false;
                }
                return controllerContext.HttpContext.Request.Form.AllKeys.Contains(this.Name);
            }
        }
        //新增申購單
        [HttpPost]
        [MultiButton("AddPR")]
        public ActionResult AddPR(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] lstQty = Request["need_qty"].Split(',');
            string[] lstDate = Request["Date_${index}"].Split(',');
            string[] lstRemark = Request["remark"].Split(',');
            //建立申購單
            log.Info("create new Purchase Requisition");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["id"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.STATUS = 10; //表示申購單已送出
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newPR(Request["id"], pr, lstItemId);
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    items.NEED_QTY = null;
                }
                else
                {
                    items.NEED_QTY = decimal.Parse(lstQty[j]);
                }
                try
                {
                    items.NEED_DATE = Convert.ToDateTime(lstDate[j]);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                }
                items.REMARK = lstRemark[j];
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.NEED_QTY + ", Date =" + items.NEED_DATE);
                lstItem.Add(items);
            }
            int k = service.refreshPR(prid, pr, lstItem);
            return Redirect("PurchaseRequisition?id=" + pr.PROJECT_ID);
        }
        [HttpPost]
        [MultiButton("SavePR")]
        //儲存申購單(申購單草稿)
        public ActionResult SavePR(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] lstQty = Request["need_qty"].Split(',');
            string[] lstDate = Request["Date_${index}"].Split(',');
            string[] lstRemark = Request["remark"].Split(',');
            //建立申購單
            log.Info("create new Purchase Requisition");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["id"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.STATUS = 0; //表示申購單未送出，只是存檔而已
            pr.PRJ_UID = int.Parse(Request["prj_uid"]);
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newPR(Request["id"], pr, lstItemId);
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    items.NEED_QTY = null;
                }
                else
                {
                    items.NEED_QTY = decimal.Parse(lstQty[j]);
                }
                try
                {
                    items.NEED_DATE = Convert.ToDateTime(lstDate[j]);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                }
                items.REMARK = lstRemark[j];
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.NEED_QTY + ", Date =" + items.NEED_DATE);
                lstItem.Add(items);
            }
            int k = service.refreshPR(prid, pr, lstItem);
            return Redirect("PurchaseRequisition?id=" + pr.PROJECT_ID);
        }
        //申購單查詢
        public ActionResult PurchaseRequisition(string id)
        {
            log.Info("Search For Purchase Requisition !!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //申購單草稿
            int status = 10;
            if (Request["status"] == null || Request["status"] == "")
            {
                status = 0;
            }
            List<PRFunction> lstPR = service.getPRByPrjId(id, Request["create_date"], Request["taskname"], Request["prid"], status);
            return View(lstPR);
        }

        public ActionResult Search()
        {
            log.Info("projectid=" + Request["id"] + ", taskname =" + Request["taskname"] + ", prid =" + Request["prid"] + ", create_id =" + Request["create_date"] + ", status =" + int.Parse(Request["status"]));
            List<PRFunction> lstPR = service.getPRByPrjId(Request["id"], Request["create_date"], Request["taskname"], Request["prid"], int.Parse(Request["status"]));
            ViewBag.SearchResult = "共取得" + lstPR.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("PurchaseRequisition", lstPR);
        }

        //顯示單一申購單功能
        public ActionResult SinglePR(string id)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(id);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            return View(singleForm);
        }

        //更新申購單草稿
        public String RefreshPR(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得申購單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            pr.PRJ_UID = int.Parse(form.Get("prjuid").Trim());
            pr.STATUS = int.Parse(form.Get("status").Trim());
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("apply_date"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("need_qty").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            string[] lstDate = form.Get("date").Split(',');
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.NEED_QTY = null;
                }
                else
                {
                    item.NEED_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Need Qty =" + item.NEED_QTY);
                item.REMARK = lstRemark[j];
                item.NEED_DATE = DateTime.ParseExact(lstDate[j], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                lstItem.Add(item);
            }
            int i = service.updatePR(formid, pr, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新申購單草稿成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid + "Task Id =" + form["prjuid"]);
            return msg;
        }

        //新增申購單
        public String CreatePR(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得申購單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            pr.PRJ_UID = int.Parse(form.Get("prjuid").Trim());
            pr.STATUS = 10;
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("apply_date"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("need_qty").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            string[] lstDate = form.Get("date").Split(',');
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.NEED_QTY = null;
                }
                else
                {
                    item.NEED_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Need Qty =" + item.NEED_QTY);
                item.REMARK = lstRemark[j];
                item.NEED_DATE = DateTime.ParseExact(lstDate[j], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                lstItem.Add(item);
            }
            int i = service.updatePR(formid, pr, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增申購單成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid + "Task Id =" + form["prjuid"]);
            return msg;
        }


        //採購作業
        public ActionResult PurchaseOrder(string id)
        {
            log.Info("Access to Purchase Order Page !!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<PurchaseOrderFunction> lstPO = service.getPRBySupplier(id);
            return View(lstPO);
        }
        //取得申購單之供應商合約項目
        public ActionResult PurchaseOperation(string id, FormCollection form)
        {
            log.Info("Access to Purchase Operation page!!");

            ViewBag.projectid = id.Substring(0, 6).Trim();
            string[] allKey = id.Split('-');
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(allKey[0]);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.supplier = allKey[2];
            ViewBag.parentPrId = allKey[1];
            ViewBag.OrderDate = DateTime.Now;
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(ViewBag.parentPrId);
            singleForm.planPR = service.formPR;
            ViewBag.recipient = singleForm.planPR.RECIPIENT;
            ViewBag.location = singleForm.planPR.LOCATION;
            ViewBag.caution = singleForm.planPR.REMARK;
            List<PurchaseRequisition> lstPR = service.getPurchaseItemBySupplier(ViewBag.parentPrId);
            return View(lstPR);
        }
        //新增採購單
        public ActionResult AddPO(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] lstQty = Request["order_qty"].Split(',');
            string[] lstPlanItemId = Request["planitemid"].Split(',');
            //建立採購單
            log.Info("create new Purchase Order");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["id"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.SUPPLIER_ID = Request["supplier"];
            pr.PARENT_PR_ID = Request["parent_pr_id"];
            pr.STATUS = 20;
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newPO(Request["id"], pr, lstItemId);
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstPlanItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    items.ORDER_QTY = null;
                }
                else
                {
                    items.ORDER_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.ORDER_QTY);
                lstItem.Add(items);
            }
            int k = service.refreshPO(prid, pr, lstItem);
            return Redirect("SinglePO?id=" + prid);
        }

        //採購單查詢
        public ActionResult PurchaseOrderIndex(string id)
        {
            log.Info("Search For Purchase Order !!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }

        [HttpPost]
        public ActionResult PurchaseOrderIndex(FormCollection f)
        {
            log.Info("projectid=" + Request["id"] + ", supplier =" + Request["supplier"] + ", prid =" + Request["prid"] + ", create_id =" + Request["create_date"]);
            List<PRFunction> lstPO = service.getPOByPrjId(Request["id"], Request["create_date"], Request["supplier"], Request["prid"]);
            ViewBag.SearchResult = "共取得" + lstPO.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("PurchaseOrderIndex", lstPO);
        }

        //顯示單一採購單功能
        public ActionResult SinglePO(string id)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(id);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            ViewBag.orderDate = singleForm.planPR.CREATE_DATE.Value.ToString("yyyy/MM/dd");
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            ViewBag.prId = service.getParentPrIdByPrId(id);
            return View(singleForm);
        }

        //更新採購單資料
        public String RefreshPO(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得採購單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.SUPPLIER_ID = form.Get("supplier").Trim();
            pr.PARENT_PR_ID = form.Get("parent_pr_id").Trim();
            pr.STATUS = int.Parse(form.Get("status").Trim());
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("order_date"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("order_qty").Split(',');

            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.ORDER_QTY = null;
                }
                else
                {
                    item.ORDER_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Order Qty =" + item.ORDER_QTY);
                lstItem.Add(item);
            }
            int i = service.updatePO(formid, pr, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新採購單成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid + " 供應商名稱=" + form["supplier"]);
            return msg;
        }
        //驗收作業
        public ActionResult Receipt(string id)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(id);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            ViewBag.receiptDate = DateTime.Now.ToString("yyyy/MM/dd");
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            return View(singleForm);
        }
        //新增驗收單資料
        public ActionResult AddReceipt(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["projectid"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] lstQty = Request["receipt_qty"].Split(',');
            string[] lstPlanItemId = Request["planitemid"].Split(',');
            //建立驗收單
            log.Info("create new Receipt");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["projectid"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.SUPPLIER_ID = Request["supplier"];
            pr.PARENT_PR_ID = Request["pr_id"];
            pr.STATUS = 30;
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newRP(Request["projectid"], pr, lstItemId);
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstPlanItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    items.RECEIPT_QTY = null;
                }
                else
                {
                    items.RECEIPT_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.RECEIPT_QTY);
                lstItem.Add(items);
            }
            int k = service.refreshRP(prid, pr, lstItem);
            return Redirect("SingleRP?id=" + prid);
        }

        //顯示單一驗收單功能
        public ActionResult SingleRP(string id)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(id);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            ViewBag.receiptDate = singleForm.planPR.CREATE_DATE.Value.ToString("yyyy/MM/dd");
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            return View(singleForm);
        }
        //更新驗收單資料
        public String RefreshRP(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得驗收單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.SUPPLIER_ID = form.Get("supplier").Trim();
            pr.PARENT_PR_ID = form.Get("parent_pr_id").Trim();
            pr.STATUS = int.Parse(form.Get("status").Trim());
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("receipt_date"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("receipt_qty").Split(',');

            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.RECEIPT_QTY = null;
                }
                else
                {
                    item.RECEIPT_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Order Qty =" + item.RECEIPT_QTY);
                lstItem.Add(item);
            }
            int i = service.updateRP(formid, pr, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新驗收單成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid);
            return msg;
        }

        //驗收單明細
        public ActionResult ReceiptList(string id)
        {
            log.Info("Access to Receipt List !!");
            List<PRFunction> lstRP = service.getRPByPrjId(id);
            return View(lstRP);
        }

        //庫存查詢
        public ActionResult InventoryIndex(string id)
        {
            log.Info("Search For Inventory of All Item !!");
            ViewBag.projectid = id;
            TnderProject tndservice = new TnderProject();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            SYS_USER u = (SYS_USER)Session["user"];
            ViewBag.createid = u.USER_ID;
            List<PurchaseRequisition> lstItem = service.getInventoryByPrjId(id, Request["item"], Request["systemMain"]);
            return View(lstItem);
        }

        public ActionResult SearchInventory()
        {
            log.Info("projectid=" + Request["id"] + ", planitemname =" + Request["item"] + ", systemMain =" + Request["systemMain"]);
            List<PurchaseRequisition> lstItem = service.getInventoryByPrjId(Request["id"], Request["item"], Request["systemMain"]);
            ViewBag.SearchResult = "共取得" + lstItem.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("InventoryIndex", lstItem);
        }

        //新增物料提領資料
        public ActionResult AddDelivery()
        {
            //取得專案編號
            log.Info("Project Id:" + Request["prjId"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["prjName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] lstQty = Request["delivery_qty"].Split(',');
            string[] lstPlanItemId = Request["planitemid"].Split(',');
            PLAN_ITEM_DELIVERY item = new PLAN_ITEM_DELIVERY();
            string deliveryorderid = service.newDelivery(Request["prjId"], lstItemId, Request["createid"]);
            List<PLAN_ITEM_DELIVERY> lstItem = new List<PLAN_ITEM_DELIVERY>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_ITEM_DELIVERY items = new PLAN_ITEM_DELIVERY();
                items.PLAN_ITEM_ID = lstPlanItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    items.DELIVERY_QTY = null;
                }
                else
                {
                    items.DELIVERY_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.DELIVERY_QTY);
                lstItem.Add(items);
            }
            int k = service.refreshDelivery(deliveryorderid, lstItem);
            return Redirect("InventoryIndex?id=" + Request["prjId"]);
        }

        //領料明細
        public ActionResult DeliveryList(string id)
        {
            log.Info("Access to Delivery List !!");
            List<PurchaseRequisition> lstItem = service.getDeliveryByItemId(id);
            PurchaseRequisition lstInventory = service.getInventoryByItemId(id);
            ViewBag.allReceipt = lstInventory.ALL_RECEIPT_QTY;
            log.Debug("plan_item_id = " + id + "其總驗收數量為" + ViewBag.allReceipt);
            return View(lstItem);
        }

        //更新領料數量
        public String RefreshDelivery(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            string[] lstItemId = form.Get("delivery_id").Split(',');
            string[] lstQty = form.Get("delivery_qty").Split(',');

            List<PLAN_ITEM_DELIVERY> lstItem = new List<PLAN_ITEM_DELIVERY>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_ITEM_DELIVERY item = new PLAN_ITEM_DELIVERY();
                item.DELIVERY_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.DELIVERY_QTY = null;
                }
                else
                {
                    item.DELIVERY_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.DELIVERY_ID + ", Delivery Qty =" + item.DELIVERY_QTY);
                lstItem.Add(item);
            }
            int i = service.updateDelivery(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新領料紀錄成功";
            }

            log.Info("Request: 更新紀錄訊息 = " + msg);
            return msg;
        }
    }
}
