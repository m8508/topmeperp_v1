using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class PlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string message = "";


        //取得標單品項資料
        public List<PLAN_ITEM> getPlanItem(string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub)
        {

            logger.Info("search plan item by 九宮格 =" + typeCode1 + "search plan item by 次九宮格 =" + typeCode2 + "search plan item by 主系統 =" + systemMain + "search plan item by 次系統 =" + systemSub);
            List<topmeperp.Models.PLAN_ITEM> lstItem = new List<PLAN_ITEM>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT * FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //九宮格
            if (null != typeCode1 && typeCode1 != "")
            {
                sql = sql + "AND p.TYPE_CODE_1 = @typeCode1 ";
                parameters.Add(new SqlParameter("typeCode1", typeCode1));
            }
            //次九宮格
            if (null != typeCode2 && typeCode2 != "")
            {
                sql = sql + "AND p.TYPE_CODE_2 = @typeCode2 ";
                parameters.Add(new SqlParameter("typeCode2", typeCode2));
            }
            //主系統
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND p.SYSTEM_MAIN LIKE @systemMain ";
                parameters.Add(new SqlParameter("systemMain", "%" + systemMain + "%"));
            }
            //次系統
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND p.SYSTEM_SUB LIKE @systemSub ";
                parameters.Add(new SqlParameter("systemSub", "%" + systemSub + "%"));
            }

            using (var context = new topmepEntities())
            {
                logger.Debug("get plan item sql=" + sql);
                lstItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get plan item count=" + lstItem.Count);
            return lstItem;
        }

        PLAN_SUP_INQUIRY form = null;
        public string newPlanForm(PLAN_SUP_INQUIRY form, string[] lstItemId)
        {
            //1.建立詢價單價單樣本
            logger.Info("create new plan form ");
            string sno_key = "PO";
            SerialKeyService snoservice = new SerialKeyService();
            form.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new plan form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_SUP_INQUIRY.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add form=" + i);
                logger.Info("plan form id = " + form.INQUIRY_FORM_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_SUP_INQUIRY_ITEM (INQUIRY_FORM_ID, INQUIRY_ITEM_ID, TYPE_CODE, "
                    + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QTY, ITEM_UNIT_PRICE, ITEM_REMARK) "
                    + "SELECT '" + form.INQUIRY_FORM_ID + "' as INQUIRY_FORM_ID, INQUIRY_ITEM_ID, TYPE_CODE_1 AS TYPE_CODE, "
                    + "TYPE_CODE_2 AS SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QUANTITY, ITEM_UNIT_PRICE, ITEM_REMARK "
                    + "FROM PLAN_ITEM where PLAN_ITEM_ID IN (" + ItemId + ")";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.INQUIRY_FORM_ID;
            }
        }
        public string getBudgetById(string prjid)
        {
            string projectid = null;
            using (var context = new topmepEntities())
            {
                projectid = context.Database.SqlQuery<string>("select DISTINCT PROJECT_ID FROM PLAN_BUDGET WHERE PROJECT_ID = @pid "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return projectid;
        }
        public int addBudget(List<PLAN_BUDGET> lstItem)
        {
            //1.新增預算資料
            int i = 0;
            logger.Info("add budget = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_BUDGET item in lstItem)
                {
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add budget count =" + i);
            return i;
        }

        public int updateBudget(string projectid, List<PLAN_BUDGET> lstItem)
        {
            //1.新增預算資料
            int i = 0;
            logger.Info("update budget = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_BUDGET item in lstItem)
                {
                    PLAN_BUDGET existItem = null;
                    logger.Debug("plan budget id=" + item.PLAN_BUDGET_ID);
                    if (item.PLAN_BUDGET_ID != 0)
                    {
                        existItem = context.PLAN_BUDGET.Find(item.PLAN_BUDGET_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("projectid", projectid));
                        parameters.Add(new SqlParameter("code1", item.TYPE_CODE_1));
                        parameters.Add(new SqlParameter("code2", item.TYPE_CODE_2));
                        string sql = "SELECT * FROM PLAN_BUDGET WHERE PROJECT_ID = @projectid and TYPE_CODE_1 + TYPE_CODE_2 = @code1 + @code2";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.TYPE_CODE_1 + item.TYPE_CODE_2);
                        PLAN_BUDGET excelItem = context.PLAN_BUDGET.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_BUDGET.Find(excelItem.PLAN_BUDGET_ID);

                    }
                    logger.Debug("find exist item=" + existItem.PLAN_BUDGET_ID);
                    existItem.BUDGET_AMOUNT = item.BUDGET_AMOUNT;
                    existItem.MODIFY_ID = item.MODIFY_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_BUDGET.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update budget count =" + i);
            return i;
        }
    }
    public class BudgetDataService : CostAnalysisDataService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public List<DirectCost> getBudget(string projectid)
        {
            List<DirectCost> lstBudget = new List<DirectCost>();
            using (var context = new topmepEntities())
            {
                string sql = "SELECT MAINCODE, MAINCODE_DESC, SUB_CODE, SUB_DESC, MATERIAL_COST, MAN_DAY,"
                    + "p.BUDGET_AMOUNT as BUDGET FROM (SELECT" +
                    "(select TYPE_CODE_1 + TYPE_CODE_2 from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE, " +
                    "(select TYPE_DESC from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE_DESC ," +
                    "(select SUB_TYPE_ID from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) T_SUB_CODE, " +
                    "TYPE_CODE_2 SUB_CODE," +
                    "(select TYPE_DESC from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) SUB_DESC," +
                    "SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) MATERIAL_COST,SUM(ITEM_QUANTITY * RATIO) MAN_DAY,count(*) ITEM_COUNT " +
                    "FROM (SELECT it.*, w.RATIO, w.PRICE FROM TND_PROJECT_ITEM it LEFT OUTER JOIN TND_WAGE w " +
                    "ON it.PROJECT_ITEM_ID = w.PROJECT_ITEM_ID WHERE it.project_id = @projectid) A " +
                    "GROUP BY TYPE_CODE_1, TYPE_CODE_2) B LEFT OUTER JOIN PLAN_BUDGET p ON MAINCODE + SUB_CODE = p.TYPE_CODE_1 + p.TYPE_CODE_2 WHERE " +
                    "p.PROJECT_ID = @projectid ORDER BY MAINCODE, SUB_CODE";
                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstBudget = context.Database.SqlQuery<DirectCost>(sql, parameters.ToArray()).ToList();
                logger.Info("Get Budget Info Record Count=" + lstBudget.Count);
            }
            return lstBudget;
        }
    }
    //採購詢價單資料提供作業
    public class PurchaseFormService : TnderProject
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public PLAN_SUP_INQUIRY formInquiry = null;
        public List<PLAN_SUP_INQUIRY_ITEM> formInquiryItem = null;

        //取得採購詢價單
        public void getInqueryForm(string formid)
        {
            logger.Info("get form : formid=" + formid);
            using (var context = new topmepEntities())
            {
                //取得詢價單檔頭資訊
                formInquiry = context.PLAN_SUP_INQUIRY.SqlQuery("SELECT * FROM PLAN_SUP_INQUIRY WHERE INQUIRY_FORM_ID=@formid", new SqlParameter("formid", formid)).First();
                //取得詢價單明細
                formInquiryItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery("SELECT * FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formid", new SqlParameter("formid", formid)).ToList();
                logger.Debug("get form item count:" + formInquiryItem.Count);
            }
        }
        //取得採購詢價單樣板(供應商欄位為0)
        public List<PLAN_SUP_INQUIRY> getFormTemplateByProject(string projectid)
        {
            logger.Info("get purchase template by projectid=" + projectid);
            List<PLAN_SUP_INQUIRY> lst = new List<PLAN_SUP_INQUIRY>();
            using (var context = new topmepEntities())
            {
                //取得詢價單樣本資訊
                lst = context.PLAN_SUP_INQUIRY.SqlQuery("SELECT * FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND　PROJECT_ID=@projectid ORDER BY INQUIRY_FORM_ID DESC",
                    new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }
        public List<PlanSupplierFormFunction> getFormByProject(string projectid)
        {
            List<PlanSupplierFormFunction> lst = new List<PlanSupplierFormFunction>();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<PlanSupplierFormFunction>("SELECT a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, SUM(b.ITEM_QTY*b.ITEM_UNIT_PRICE) AS TOTAL_PRICE, ROW_NUMBER() OVER(ORDER BY a.INQUIRY_FORM_ID DESC) AS NO FROM PLAN_SUP_INQUIRY a left JOIN PLAN_SUP_INQUIRY_ITEM b ON a.INQUIRY_FORM_ID = b.INQUIRY_FORM_ID GROUP BY a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, a.PROJECT_ID HAVING  a.SUPPLIER_ID IS NOT NULL AND a.PROJECT_ID =@projectid ORDER BY a.INQUIRY_FORM_ID DESC, a.FORM_NAME ;", new SqlParameter("projectid", projectid)).ToList();
            }
            logger.Info("get plan supplier form function count:" + lst.Count);
            return lst;
        }

        //新增採購供應商詢價單
        public string addSupplierForm(PLAN_SUP_INQUIRY sf, string[] lstItemId)
        {
            string message = "";
            string sno_key = "PO";
            SerialKeyService snoservice = new SerialKeyService();
            sf.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_SUP_INQUIRY.AddOrUpdate(sf);
                    i = context.SaveChanges();
                    List<topmeperp.Models.PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
                    string ItemId = "";
                    for (i = 0; i < lstItemId.Count(); i++)
                    {
                        if (i < lstItemId.Count() - 1)
                        {
                            ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                        }
                        else
                        {
                            ItemId = ItemId + "'" + lstItemId[i] + "'";
                        }
                    }

                    string sql = "INSERT INTO PLAN_SUP_INQUIRY_ITEM (INQUIRY_FORM_ID, PLAN_ITEM_ID,"
                        + "TYPE_CODE, SUB_TYPE_CODE,ITEM_DESC,ITEM_UNIT, ITEM_QTY,"
                        + "ITEM_UNIT_PRICE, ITEM_REMARK) "
                        + "SELECT '" + sf.INQUIRY_FORM_ID + "' as INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE,"
                        + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT,ITEM_QTY, ITEM_UNIT_PRICE,"
                        + "ITEM_REMARK "
                        + "FROM PLAN_SUP_INQUIRY_ITEM where INQUIRY_FORM_ITEM_ID IN (" + ItemId + ")";

                    logger.Info("sql =" + sql);
                    var parameters = new List<SqlParameter>();
                    i = context.Database.ExecuteSqlCommand(sql);

                }
                catch (Exception e)
                {
                    logger.Error("add new plan supplier form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return sf.INQUIRY_FORM_ID;
        }
        PLAN_SUP_INQUIRY form = null;
        //更新採購廠商詢價單資料
        public int refreshPlanSupplierForm(string formid, PLAN_SUP_INQUIRY sf, List<PLAN_SUP_INQUIRY_ITEM> lstItem)
        {
            logger.Info("Update plan supplier inquiry form id =" + formid);
            form = sf;
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan supplier inquiry form =" + i);
                    logger.Info("supplier inquiry form item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_SUP_INQUIRY_ITEM item in lstItem)
                    {
                        PLAN_SUP_INQUIRY_ITEM existItem = null;
                        logger.Debug("form item id=" + item.INQUIRY_ITEM_ID);
                        if (item.INQUIRY_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_SUP_INQUIRY_ITEM.Find(item.INQUIRY_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                            PLAN_SUP_INQUIRY_ITEM excelItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_SUP_INQUIRY_ITEM.Find(excelItem.INQUIRY_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.ITEM_DESC);
                        existItem.ITEM_UNIT_PRICE = item.ITEM_UNIT_PRICE;
                        context.PLAN_SUP_INQUIRY_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update plan supplier inquiry form item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new plan supplier form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        public int createPlanFormFromSupplier(PLAN_SUP_INQUIRY form, List<PLAN_SUP_INQUIRY_ITEM> items)
        {
            int i = 0;
            //1.建立詢價單價單樣本
            string sno_key = "PO";
            SerialKeyService snoservice = new SerialKeyService();
            form.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("Plan form from supplier =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_SUP_INQUIRY.Add(form);

                logger.Info("plan form id = " + form.INQUIRY_FORM_ID);
                //if (i > 0) { status = true; };
                foreach (PLAN_SUP_INQUIRY_ITEM item in items)
                {
                    item.INQUIRY_FORM_ID = form.INQUIRY_FORM_ID;
                    context.PLAN_SUP_INQUIRY_ITEM.Add(item);
                }
                i = context.SaveChanges();
            }
            return i;
        }
        public List<string> getSystemMain(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SYSTEM_MAIN FROM PLAN_ITEM　WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get System Main Count=" + lst.Count);
            }
            return lst;
        }
        //取得供應商選單
        public List<string> getSupplier()
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得供應商選單
                lst = context.Database.SqlQuery<string>("SELECT (SELECT SUPPLIER_ID + '' + COMPANY_NAME FROM TND_SUPPLIER s2 WHERE s2.SUPPLIER_ID = s1.SUPPLIER_ID for XML PATH('')) AS suppliers FROM TND_SUPPLIER s1 ;").ToList();
                logger.Info("Get Supplier Count=" + lst.Count);
            }
            return lst;
        }
        //取得次系統選單
        public List<string> getSystemSub(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SYSTEM_SUB FROM PLAN_ITEM WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                //lst = context.TND_PROJECT_ITEM.SqlQuery("SELECT DISTINCT SYSTEM_SUB FROM TND_PROJECT_ITEM　WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get System Sub Count=" + lst.Count);
            }
            return lst;
        }

    }
}
