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
        TND_PROJECT project = null;
        public TND_PROJECT budgetTable = null;
        public List<PLAN_BUDGET> budgetTableItem = null;


        #region 得標標單項目處理
        public TND_PROJECT getProject(string prjid)
        {
            using (var context = new topmepEntities())
            {
                project = context.TND_PROJECT.SqlQuery("select p.* from TND_PROJECT p "
                    + "where p.PROJECT_ID = @pid "
                   , new SqlParameter("pid", prjid)).First();
            }
            return project;
        }
        public int delAllItem()
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all item by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_ITEM WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", project.PROJECT_ID));
            }
            logger.Debug("delete item count=" + i);
            return i;
        }
        public int refreshItem(List<PLAN_ITEM> planItem)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshPlanItem = " + planItem.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_ITEM item in planItem)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.PLAN_ITEM.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add plan item count =" + i);
            return i;
        }
        #endregion

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

        budgetsummary budget = null;
        public budgetsummary getBudgetForComparison(string projectid, string typecode1, string typecode2, string systemMain, string systemSub)
        {
            using (var context = new topmepEntities())
            {
                if (null != typecode1 && typecode1 != "" && typecode2 == "" || null != typecode1 && typecode1 != "" && typecode2 == null)
                {
                    budget = context.Database.SqlQuery<budgetsummary>("SELECT TYPE_CODE_1,SUM(BUDGET_AMOUNT) AS BAmount " +
                        "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid GROUP BY TYPE_CODE_1 HAVING TYPE_CODE_1 = @typecode1 "
                       , new SqlParameter("pid", projectid), new SqlParameter("typecode1", typecode1)).First();
                }
                else if (null != typecode1 && typecode1 != "" && null != typecode2 && typecode2 != "")
                {
                    budget = context.Database.SqlQuery<budgetsummary>("SELECT TYPE_CODE_1, TYPE_CODE_1, BUDGET_AMOUNT AS BAmount " +
                        "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid AND TYPE_CODE_1 = @typecode1 AND TYPE_CODE_2 = @typecode2 "
                       , new SqlParameter("pid", projectid), new SqlParameter("typecode1", typecode1), new SqlParameter("typecode2", typecode2)).FirstOrDefault();
                }
                else if (null == typecode1 && null == typecode2 && null == systemMain && null == systemSub || typecode1 == "" && typecode2 == "" && systemMain == "" && systemSub == "")
                {
                    budget = context.Database.SqlQuery<budgetsummary>("SELECT SUM(BUDGET_AMOUNT) AS BAmount " +
                        "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid "
                       , new SqlParameter("pid", projectid)).First();
                }
                else
                {
                    budget = null;
                }
            }
            return budget;
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
            //1.修改預算資料
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
                    existItem.BUDGET_RATIO = item.BUDGET_RATIO;
                    existItem.MODIFY_ID = item.MODIFY_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_BUDGET.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update budget count =" + i);
            return i;
        }
        public int updateBudgetToPlanItem(string id)
        {
            int i = 0;
            logger.Info("update budget ratio to plan items by id :" + id);
            string sql = "UPDATE PLAN_ITEM SET PLAN_ITEM.BUDGET_RATIO = plan_budget.BUDGET_RATIO, " +
                   "PLAN_ITEM.TND_RATIO = plan_budget.TND_RATIO from PLAN_ITEM inner join " +
                   "plan_budget on @id + PLAN_ITEM.TYPE_CODE_1 + PLAN_ITEM.TYPE_CODE_2 " +
                   "= @id + plan_budget.TYPE_CODE_1 + plan_budget.TYPE_CODE_2 ";
            logger.Debug("sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("id", id));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //更新得標標單品項個別預算
        public int updateItemBudget(string projectid, List<PLAN_ITEM> lstItem)
        {
            //1.新增預算資料
            int i = 0;
            logger.Info("update budget = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_ITEM item in lstItem)
                {
                    PLAN_ITEM existItem = null;
                    logger.Debug("plan item id=" + item.PLAN_ITEM_ID);
                    if (item.PLAN_ITEM_ID != null && item.PLAN_ITEM_ID != "")
                    {
                        existItem = context.PLAN_ITEM.Find(item.PLAN_ITEM_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("projectid", projectid));
                        parameters.Add(new SqlParameter("planitemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ITEM WHERE PROJECT_ID = @projectid and PLAN_ITEM_ID = @planitemid ";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.PLAN_ITEM_ID);
                        PLAN_ITEM excelItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ITEM.Find(excelItem.PLAN_ITEM_ID);

                    }
                    logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                    existItem.BUDGET_RATIO = item.BUDGET_RATIO;
                    existItem.MODIFY_USER_ID = item.MODIFY_USER_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_ITEM.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update budget count =" + i);
            return i;
        }

        #region 取得預算表單檔頭資訊
        //取得預算表單檔頭
        public void getProjectId(string projectid)
        {
            logger.Info("get project : projectid=" + projectid);
            using (var context = new topmepEntities())
            {
                //取得預算表單檔頭資訊
                budgetTable = context.TND_PROJECT.SqlQuery("SELECT * FROM TND_PROJECT WHERE PROJECT_ID=@projectid", new SqlParameter("projectid", projectid)).First();
            }
        }
        #endregion

        #region 預算數量  
        //預算上傳數量  
        public int refreshBudget(List<PLAN_BUDGET> items)
        {
            //1.檢查專案是否存在
            //if (null == project) { throw new Exception("Project is not exist !!"); } 先註解掉,因為讀取不到project,會造成null == project is true,
            //而導致錯誤, 因為已設定是直接由專案頁面導入上傳圖算畫面，故不會有專案不存在的bug
            int i = 0;
            logger.Info("refreshBudgetItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_BUDGET item in items)
                {
                    //item.PROJECT_ID = project.PROJECT_ID;先註解掉,因為專案編號一開始已經設定了，會直接代入
                    context.PLAN_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add PLAN_BUDGET count =" + i);
            return i;
        }
        public int delBudgetByProject(string projectid)
        {
            logger.Info("remove all budget by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all PLAN_BUDGET by proejct id=" + projectid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_BUDGET WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete PLAN_BUDGET count=" + i);
            return i;
        }
        #endregion

        PlanRevenue plan = null;
        public PlanRevenue getPlanRevenueById(string prjid)
        {
            using (var context = new topmepEntities())
            {
                plan = context.Database.SqlQuery<PlanRevenue>("SELECT p.PROJECT_ID + p.PROJECT_NAME AS CONTRACT_ID, " +
                    "(SELECT SUM(ITEM_UNIT_COST*ITEM_QUANTITY) FROM PLAN_ITEM pi WHERE pi.PROJECT_ID = @pid) AS PLAN_REVENUE " +
                     "FROM TND_PROJECT p WHERE p.PROJECT_ID = @pid "
                   , new SqlParameter("pid", prjid)).First();
            }
            return plan;
        }

        public int addContractId4Owner(string projectid)
        {
            int i = 0;
            //將業主合約編號寫入PLAN PAYMENT TERMS
            logger.Info("copy contract id from owner into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                PLAN_PAYMENT_TERMS lstItem = new PLAN_PAYMENT_TERMS();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID) " +
                       "SELECT '" + projectid + "'  + p.PROJECT_NAME AS contractid, '" + projectid + "'  FROM TND_PROJECT p WHERE p.PROJECT_ID = '" + projectid + "'  " +
                       "AND '" + projectid + "'  + p.PROJECT_NAME NOT IN(SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
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
                string sql = "SELECT C.*, SUM((MATERIAL_COST * COST_RATIO / 100 + MAN_DAY) * BUDGET / 100) AS AMOUNT_BY_CODE FROM " +
                    "(SELECT MAINCODE, MAINCODE_DESC, SUB_CODE, SUB_DESC, MATERIAL_COST, MAN_DAY, "
                    + "BUDGET_RATIO as BUDGET, COST_RATIO FROM (SELECT" +
                    "(select TYPE_CODE_1 + TYPE_CODE_2 from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE, " +
                    "(select TYPE_DESC from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE_DESC ," +
                    "(select SUB_TYPE_ID from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) T_SUB_CODE, " +
                    "TYPE_CODE_2 SUB_CODE," +
                    "(select TYPE_DESC from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) SUB_DESC," +
                    "SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) MATERIAL_COST, SUM(ITEM_QUANTITY * PRICE) MAN_DAY,count(*) ITEM_COUNT " +
                    "FROM (SELECT it.*, w.RATIO, w.PRICE FROM TND_PROJECT_ITEM it LEFT OUTER JOIN TND_WAGE w " +
                    "ON it.PROJECT_ITEM_ID = w.PROJECT_ITEM_ID WHERE it.project_id = @projectid) A " +
                    "GROUP BY TYPE_CODE_1, TYPE_CODE_2) B LEFT OUTER JOIN (SELECT p.TYPE_CODE_1, p.TYPE_CODE_2, SUM(p.BUDGET_RATIO*p.ITEM_QUANTITY)/SUM(p.ITEM_QUANTITY) BUDGET_RATIO, " +
                    "SUM(p.TND_RATIO*p.ITEM_QUANTITY)/SUM(p.ITEM_QUANTITY) COST_RATIO FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid GROUP BY p.TYPE_CODE_1, p.TYPE_CODE_2 ) D ON MAINCODE + SUB_CODE = D.TYPE_CODE_1 + D.TYPE_CODE_2 " +
                    ") C GROUP BY MAINCODE, MAINCODE_DESC, SUB_CODE, SUB_DESC, MATERIAL_COST, MAN_DAY, BUDGET, COST_RATIO ORDER BY MAINCODE, SUB_CODE";
                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstBudget = context.Database.SqlQuery<DirectCost>(sql, parameters.ToArray()).ToList();
                logger.Info("Get Budget Info Record Count=" + lstBudget.Count);
            }
            return lstBudget;
        }
        //取得投標標單總直接成本
        public DirectCost getTotalCost(string projectid)
        {
            DirectCost lstTotalCost = null;
            using (var context = new topmepEntities())
            {
                string sql = "SELECT SUM(TND_COST) AS TOTAL_COST, SUM(BUDGET) AS TOTAL_BUDGET, SUM(P_COST) AS TOTAL_P_COST FROM (SELECT(select TYPE_CODE_1 + TYPE_CODE_2 from REF_TYPE_MAIN WHERE  " +
                    "TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE, (select TYPE_DESC from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE_DESC, " +
                    "(select SUB_TYPE_ID from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) T_SUB_CODE, TYPE_CODE_2 SUB_CODE, " +
                    "(select TYPE_DESC from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) SUB_DESC, (SUM(ITEM_QUANTITY * ITEM_UNIT_COST)- SUM(ITEM_QUANTITY * PRICE)) MATERIAL_COST, " +
                    "SUM(ITEM_QUANTITY * PRICE) MAN_DAY,count(*) ITEM_COUNT, SUM(ITEM_QUANTITY * ITEM_UNIT_COST * BUDGET_RATIO/100) BUDGET, " +
                    "SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) + SUM(ITEM_QUANTITY * MAN_PRICE) P_COST, SUM(ITEM_QUANTITY * ITEM_UNIT_COST) TND_COST FROM " +
                    "(SELECT it.*, w.RATIO, w.PRICE FROM PLAN_ITEM it LEFT OUTER JOIN TND_WAGE w ON it.PLAN_ITEM_ID = w.PROJECT_ITEM_ID WHERE it.project_id = @projectid) A  " +
                    "GROUP BY TYPE_CODE_1, TYPE_CODE_2)B ";
                logger.Info("sql = " + sql);
                lstTotalCost = context.Database.SqlQuery<DirectCost>(sql, new SqlParameter("projectid", projectid)).First();
            }
            return lstTotalCost;
        }

        #region 取得特定標單項目材料成本與預算
        //取得特定標單項目材料成本與預算
        public DirectCost getItemBudget(string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub, string formName)
        {
            logger.Info("search plan item by 九宮格 =" + typeCode1 + "search plan item by 次九宮格 =" + typeCode2 + "search plan item by 主系統 =" + systemMain + "search plan item by 次系統 =" + systemSub + "search plan item by 採購名稱 =" + formName);
            DirectCost lstItemBudget = null;
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT SUM(pi.ITEM_QUANTITY*pi.ITEM_UNIT_COST) AS ITEM_COST, SUM(pi.ITEM_QUANTITY*pi.ITEM_UNIT_COST*pi.BUDGET_RATIO/100) AS ITEM_BUDGET FROM PLAN_ITEM pi ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //採購項目
            if (null != formName && formName != "")
            {
                sql = sql + "right join (select distinct p.FORM_NAME + pii.PLAN_ITEM_ID as FORM_KEY, p.FORM_NAME, pii.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY p " +
                    "LEFT JOIN PLAN_SUP_INQUIRY_ITEM pii on p.INQUIRY_FORM_ID = pii.INQUIRY_FORM_ID WHERE p.FORM_NAME =@formName)A ON pi.PLAN_ITEM_ID = A.PLAN_ITEM_ID ";
                parameters.Add(new SqlParameter("formName", formName));
            }
            sql = sql + " WHERE pi.PROJECT_ID =@projectid ";
            //九宮格
            if (null != typeCode1 && typeCode1 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 = @typeCode1 ";
                parameters.Add(new SqlParameter("typeCode1", typeCode1));
            }
            //次九宮格
            if (null != typeCode2 && typeCode2 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_2 = @typeCode2 ";
                parameters.Add(new SqlParameter("typeCode2", typeCode2));
            }
            //主系統
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND pi.SYSTEM_MAIN LIKE @systemMain ";
                parameters.Add(new SqlParameter("systemMain", "%" + systemMain + "%"));
            }
            //次系統
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND pi.SYSTEM_SUB LIKE @systemSub ";
                parameters.Add(new SqlParameter("systemSub", "%" + systemSub + "%"));
            }

            using (var context = new topmepEntities())
            {
                logger.Debug("get plan item sql=" + sql);
                lstItemBudget = context.Database.SqlQuery<DirectCost>(sql, parameters.ToArray()).First();
            }
            return lstItemBudget;
        }
        #endregion
        
    }
    //採購詢價單資料提供作業
    public class PurchaseFormService : TnderProject
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public PLAN_SUP_INQUIRY formInquiry = null;
        public List<PLAN_SUP_INQUIRY_ITEM> formInquiryItem = null;
        public Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = null;

        #region 取得得標標單項目內容
        //取得標單品項資料
        public List<PLAN_ITEM> getPlanItem(string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub, string formName, string supplier)
        {

            logger.Info("search plan item by 九宮格 =" + typeCode1 + "search plan item by 次九宮格 =" + typeCode2 + "search plan item by 主系統 =" + systemMain + "search plan item by 次系統 =" + systemSub + "search plan item by 採購項目 =" + formName + "search plan item by 材料供應商 =" + supplier);
            List<topmeperp.Models.PLAN_ITEM> lstItem = new List<PLAN_ITEM>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT * FROM PLAN_ITEM pi ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //採購項目
            if (null != formName && formName != "")
            {
                sql = sql + "right join (select distinct p.FORM_NAME + pii.PLAN_ITEM_ID as FORM_KEY, p.FORM_NAME, pii.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY p " +
                    "LEFT JOIN PLAN_SUP_INQUIRY_ITEM pii on p.INQUIRY_FORM_ID = pii.INQUIRY_FORM_ID WHERE p.FORM_NAME =@formName)A ON pi.PLAN_ITEM_ID = A.PLAN_ITEM_ID ";
                parameters.Add(new SqlParameter("formName", formName));
            }
            sql = sql + " WHERE pi.PROJECT_ID =@projectid ";
            //九宮格
            if (null != typeCode1 && typeCode1 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 = @typeCode1 ";
                parameters.Add(new SqlParameter("typeCode1", typeCode1));
            }
            //次九宮格
            if (null != typeCode2 && typeCode2 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_2 = @typeCode2 ";
                parameters.Add(new SqlParameter("typeCode2", typeCode2));
            }
            //主系統
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND pi.SYSTEM_MAIN LIKE @systemMain ";
                parameters.Add(new SqlParameter("systemMain", "%" + systemMain + "%"));
            }
            //次系統
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND pi.SYSTEM_SUB LIKE @systemSub ";
                parameters.Add(new SqlParameter("systemSub", "%" + systemSub + "%"));
            }
            //材料供應商
            if (null != supplier && supplier != "")
            {
                sql = sql + "AND pi.SUPPLIER_ID =@supplier ";
                parameters.Add(new SqlParameter("supplier", supplier));
            }

            using (var context = new topmepEntities())
            {
                logger.Debug("get plan item sql=" + sql);
                lstItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get plan item count=" + lstItem.Count);
            return lstItem;
        }
        #endregion

        public PLAN_ITEM getPlanItem(string itemid)
        {
            logger.Debug("get plan item by id=" + itemid);
            PLAN_ITEM pitem = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                pitem = context.PLAN_ITEM.SqlQuery("SELECT * FROM PLAN_ITEM WHERE PLAN_ITEM_ID=@itemid",
                new SqlParameter("itemid", itemid)).First();
            }
            return pitem;
        }

        public int updatePlanItem(PLAN_ITEM item)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_ITEM.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("updatePlanItem  fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //批次產生空白表單
        public int createPlanEmptyForm(string projectid, SYS_USER loginUser)
        {
            int i = 0;
            int i2 = 0;
            using (var context = new topmepEntities())
            {
                //0.清除所有空白詢價單樣板
                string sql = "DELETE FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID IN (SELECT INQUIRY_FORM_ID FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID=@projectid);";
                i2 = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", projectid));
                logger.Info("delete template inquiry form item  by porjectid=" + projectid + ",result=" + i2);
                sql = "DELETE FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID=@projectid; ";
                i2 = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", projectid));
                logger.Info("delete template inquiry form  by porjectid=" + projectid + ",result=" + i2);

                //1.依據專案取得九宮格次九宮格分類.
                sql = "SELECT DISTINCT isnull(TYPE_CODE_1,'未分類') TYPE_CODE_1," +
                   "(SELECT TYPE_DESC FROM REF_TYPE_MAIN m WHERE m.TYPE_CODE_1 + m.TYPE_CODE_2 = p.TYPE_CODE_1) as TYPE_CODE_1_NAME, " +
                   "isnull(TYPE_CODE_2,'未分類') TYPE_CODE_2," +
                   "(SELECT TYPE_DESC FROM REF_TYPE_SUB sub WHERE sub.TYPE_CODE_ID = p.TYPE_CODE_1 AND sub.SUB_TYPE_CODE = p.TYPE_CODE_2) as TYPE_CODE_2_NAME " +
                   "FROM TND_PROJECT_ITEM p WHERE PROJECT_ID = @projectid ORDER BY TYPE_CODE_1 ,Type_CODE_2; ";

                List<TYPE_CODE_INDEX> lstType = context.Database.SqlQuery<TYPE_CODE_INDEX>(sql, new SqlParameter("projectid", projectid)).ToList();
                logger.Debug("get type index count=" + lstType.Count);
                foreach (TYPE_CODE_INDEX idx in lstType)
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("projectid", projectid));
                    sql = "SELECT * FROM PLAN_ITEM WHERE PROJECT_ID = @projectid ";
                    if (idx.TYPE_CODE_1 == "未分類")
                    {
                        sql = sql + "AND TYPE_CODE_1 is null ";
                    }
                    else
                    {
                        sql = sql + "AND TYPE_CODE_1=@typecode1 ";
                        parameters.Add(new SqlParameter("typecode1", idx.TYPE_CODE_1));
                    }

                    if (idx.TYPE_CODE_2 == "未分類")
                    {
                        sql = sql + "AND TYPE_CODE_2 is null ";
                    }
                    else
                    {
                        sql = sql + "AND TYPE_CODE_2=@typecode2 ";
                        parameters.Add(new SqlParameter("typecode2", idx.TYPE_CODE_2));
                    }
                    //2.依據分類取得詢價單項次
                    List<PLAN_ITEM> lstPlanItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
                    logger.Debug("get plan item count=" + lstPlanItem.Count + ", by typecode1=" + idx.TYPE_CODE_1 + ",typeCode2=" + idx.TYPE_CODE_2);
                    string[] itemId = new string[lstPlanItem.Count];
                    int j = 0;
                    foreach (PLAN_ITEM item in lstPlanItem)
                    {
                        itemId[j] = item.PLAN_ITEM_ID;
                        j++;
                    }
                    //3.建立詢價單基本資料
                    PLAN_SUP_INQUIRY f = new PLAN_SUP_INQUIRY();
                    if (idx.TYPE_CODE_1 == "未分類")
                    {
                        f.FORM_NAME = "未分類";
                    }
                    else
                    {
                        f.FORM_NAME = idx.TYPE_CODE_1_NAME;
                    }

                    if (idx.TYPE_CODE_2 != "未分類")
                    {
                        f.FORM_NAME = f.FORM_NAME + "-" + idx.TYPE_CODE_2_NAME;
                    }
                    f.FORM_NAME = f.FORM_NAME + "(" + idx.TYPE_CODE_1 + "," + idx.TYPE_CODE_2 + ")";
                    f.PROJECT_ID = projectid;
                    f.CREATE_ID = loginUser.USER_ID;
                    f.CREATE_DATE = DateTime.Now;
                    f.OWNER_NAME = loginUser.USER_NAME;
                    f.OWNER_EMAIL = loginUser.EMAIL;
                    f.OWNER_TEL = loginUser.TEL;
                    f.OWNER_FAX = loginUser.FAX;
                    //4.建立表單
                    string fid = newPlanForm(f, itemId);
                    logger.Info("create template form:" + fid);
                    i++;
                }
            }
            logger.Info("create form count" + i);
            return i;
        }

        public string newPlanForm(PLAN_SUP_INQUIRY form, string[] lstItemId)
        {
            //1.建立詢價單價單樣本
            logger.Info("create new plan form ");
            string sno_key = "PP";
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

                string sql = "INSERT INTO PLAN_SUP_INQUIRY_ITEM (INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE, "
                    + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QTY, ITEM_UNIT_PRICE, ITEM_REMARK) "
                    + "SELECT '" + form.INQUIRY_FORM_ID + "' as INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE_1 AS TYPE_CODE, "
                    + "TYPE_CODE_2 AS SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QUANTITY, ITEM_UNIT_PRICE, ITEM_REMARK "
                    + "FROM PLAN_ITEM where PLAN_ITEM_ID IN (" + ItemId + ")";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.INQUIRY_FORM_ID;
            }
        }
        //取得採購詢價單
        public void getInqueryForm(string formid)
        {
            logger.Info("get form : formid=" + formid);
            using (var context = new topmepEntities())
            {
                //取得詢價單檔頭資訊
                string sql = "SELECT INQUIRY_FORM_ID,PROJECT_ID,FORM_NAME,OWNER_NAME,OWNER_TEL "
                    + ",OWNER_EMAIL, OWNER_FAX, SUPPLIER_ID, CONTACT_NAME, CONTACT_EMAIL "
                    + ",DUEDATE, REF_ID, CREATE_ID, CREATE_DATE, MODIFY_ID"
                    + ",MODIFY_DATE,ISNULL(STATUS,'有效') as STATUS,ISNULL(ISWAGE,'N') as ISWAGE "
                    + "FROM PLAN_SUP_INQUIRY WHERE INQUIRY_FORM_ID = @formid";
                formInquiry = context.PLAN_SUP_INQUIRY.SqlQuery(sql, new SqlParameter("formid", formid)).First();
                //取得詢價單明細
                formInquiryItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery("SELECT * FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formid", new SqlParameter("formid", formid)).ToList();
                logger.Debug("get form item count:" + formInquiryItem.Count);
            }
        }
        int i = 0;
        // 取得採購詢價單預算金額
        public List<COMPARASION_DATA_4PLAN> getBudgetForComparison(string projectid, string formname)
        {
            List<COMPARASION_DATA_4PLAN> budget = new List<COMPARASION_DATA_4PLAN>();
            string[] eachname = formname.Split(',');
            string ItemId = "";
            for (i = 0; i < eachname.Count(); i++)
            {
                if (i < eachname.Count() - 1)
                {
                    ItemId = ItemId + "'" + eachname[i] + "'" + ",";
                }
                else
                {
                    ItemId = ItemId + "'" + eachname[i] + "'";
                }
            }
            string sql = "SELECT FORM_NAME, BUDGET_AMOUNT AS BAmount " +
                "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid AND FORM_NAME IN (" + ItemId + ") ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("pid", projectid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get sql=" + sql);
                budget = context.Database.SqlQuery<COMPARASION_DATA_4PLAN>(sql, parameters.ToArray()).ToList();
            }
            return budget;
        }
        public string zipAllTemplate4Download(string projectid)
        {
            //1.取得專案所有空白詢價單
            List<PLAN_SUP_INQUIRY> lstTemplate = getFormTemplateByProject(projectid);
            ZipFileCreator zipTool = new ZipFileCreator();
            //2.設定暫存目錄
            string tempFolder = ContextService.strUploadPath + "\\" + projectid + "\\" + ContextService.quotesFolder + "\\Temp\\";
            ZipFileCreator.DelDirectory(tempFolder);
            ZipFileCreator.CreateDirectory(tempFolder);
            //3.批次產生空白詢價單
            PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
            TND_PROJECT p = getProjectById(projectid);
            foreach (PLAN_SUP_INQUIRY f in lstTemplate)
            {
                getInqueryForm(f.INQUIRY_FORM_ID);
                string fileLocation = poi.exportExcel4po(formInquiry, formInquiryItem, true);
                logger.Debug("temp file=" + fileLocation);
            }
            //4.Zip all file
            return zipTool.ZipFiles(tempFolder, null, p.PROJECT_NAME);
        }
        //取得採購詢價單樣板(供應商欄位為0)
        public List<PLAN_SUP_INQUIRY> getFormTemplateByProject(string projectid)
        {
            logger.Info("get purchase template by projectid=" + projectid);
            List<PLAN_SUP_INQUIRY> lst = new List<PLAN_SUP_INQUIRY>();
            using (var context = new topmepEntities())
            {
                //取得詢價單樣本資訊
                string sql = "SELECT INQUIRY_FORM_ID, PROJECT_ID,FORM_NAME,OWNER_NAME,OWNER_TEL,OWNER_EMAIL "
                    + ",OWNER_FAX,SUPPLIER_ID,CONTACT_NAME,CONTACT_EMAIL,DUEDATE,REF_ID,CREATE_ID,CREATE_DATE "
                    + ",MODIFY_ID,MODIFY_DATE,ISNULL(STATUS,'有效') STATUS, ISNULL(ISWAGE,'N') ISWAGE "
                    + "FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID =@projectid ORDER BY INQUIRY_FORM_ID DESC";
                lst = context.PLAN_SUP_INQUIRY.SqlQuery(sql, new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }



        public List<PlanSupplierFormFunction> getFormByProject(string projectid, string _status)
        {
            string status = "有效";
            if (null != _status && _status != "*")
            {
                status = _status;
            }
            List<PlanSupplierFormFunction> lst = new List<PlanSupplierFormFunction>();
            using (var context = new topmepEntities())
            {
                string sql = "SELECT a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, SUM(b.ITEM_QTY*b.ITEM_UNIT_PRICE) AS TOTAL_PRICE, ROW_NUMBER() OVER(ORDER BY a.INQUIRY_FORM_ID DESC) AS NO, ISNULL(A.STATUS, '有效') AS STATUS, ISNULL(A.ISWAGE,'N') ISWAGE " +
                    "FROM PLAN_SUP_INQUIRY a left JOIN PLAN_SUP_INQUIRY_ITEM b ON a.INQUIRY_FORM_ID = b.INQUIRY_FORM_ID WHERE ISNULL(A.STATUS,'有效')=@status GROUP BY a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, a.PROJECT_ID, a.STATUS, a.ISWAGE HAVING  a.SUPPLIER_ID IS NOT NULL " +
                    "AND a.PROJECT_ID =@projectid ORDER BY a.INQUIRY_FORM_ID DESC, a.FORM_NAME ";
                lst = context.Database.SqlQuery<PlanSupplierFormFunction>(sql, new SqlParameter("status", status), new SqlParameter("projectid", projectid)).ToList();
            }
            logger.Info("get plan supplier form function count:" + lst.Count);
            return lst;
        }

        public int addFormName(List<PLAN_SUP_INQUIRY> lstItem)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Info(" No. of plan form to refresh  = " + lstItem.Count);
                    //2.將plan form資料寫入 
                    foreach (PLAN_SUP_INQUIRY item in lstItem)
                    {
                        PLAN_SUP_INQUIRY existItem = null;
                        logger.Debug("plan form id=" + item.INQUIRY_FORM_ID);
                        if (item.INQUIRY_FORM_ID != null)
                        {
                            existItem = context.PLAN_SUP_INQUIRY.Find(item.INQUIRY_FORM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", item.INQUIRY_FORM_ID));
                            string sql = "SELECT * FROM PLAN_SUP_INQUIRY WHERE INQUIRY_FORM_ID=@formid";
                            logger.Info(sql + " ;" + item.INQUIRY_FORM_ID);
                            PLAN_SUP_INQUIRY excelItem = context.PLAN_SUP_INQUIRY.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_SUP_INQUIRY.Find(excelItem.INQUIRY_FORM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PROJECT_ID + " ;" + existItem.INQUIRY_FORM_ID);
                        existItem.FORM_NAME = item.FORM_NAME;
                        context.PLAN_SUP_INQUIRY.AddOrUpdate(existItem);
                    }
                    i = context.SaveChanges();
                    logger.Debug("No. of update plan form =" + i);
                    return i;

                }
                catch (Exception e)
                {
                    logger.Error("update  plan  form  fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }

        //新增供應商採購詢價單
        public string addSupplierForm(PLAN_SUP_INQUIRY sf, string[] lstItemId)
        {
            string message = "";
            string sno_key = "PP";
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
                        + "FROM PLAN_SUP_INQUIRY_ITEM where INQUIRY_ITEM_ID IN (" + ItemId + ")";

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
        //更新供應商採購詢價單資料
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
            string sno_key = "PP";
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
        //取得材料合約供應商選單
        public List<string> getSupplierForContract(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得供應商選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SUPPLIER_ID FROM PLAN_ITEM WHERE PROJECT_ID=@projectid AND SUPPLIER_ID IS NOT NULL ;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get Supplier For Contract Count=" + lst.Count);
            }
            return lst;
        }
        //取得材料合約採購項目名稱
        public List<string> getFormNameForContract(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得供應商選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT FORM_NAME FROM PLAN_ITEM WHERE PROJECT_ID=@projectid AND SUPPLIER_ID IS NOT NULL ;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get Supplier For Contract Count=" + lst.Count);
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

        //取得個別材料廠商合約資料與金額
        public List<plansummary> getPlanContract(string projectid)
        {
            List<plansummary> lst = new List<plansummary>();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT  p.PROJECT_ID + p.SUPPLIER_ID + p.FORM_NAME AS CONTRACT_ID, p.SUPPLIER_ID, p.FORM_NAME, " +
                    "SUM(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) MATERIAL_COST, SUM(p.ITEM_QUANTITY * ISNULL(p.MAN_PRICE,0)) WAGE_COST, " +
                    "SUM(p.ITEM_QUANTITY * p.ITEM_UNIT_COST) REVENUE, SUM(p.ITEM_QUANTITY * p.ITEM_UNIT_COST * p.BUDGET_RATIO / 100) BUDGET, " +
                    "(SUM(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) + SUM(p.ITEM_QUANTITY * ISNULL(p.MAN_PRICE,0))) COST, (SUM(p.ITEM_QUANTITY * p.ITEM_UNIT_COST) - " +
                    "SUM(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) - SUM(p.ITEM_QUANTITY * ISNULL(p.MAN_PRICE,0))) PROFIT, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY p.SUPPLIER_ID) AS NO FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid and p.ITEM_UNIT_PRICE IS NOT NULL " +
                    "AND p.ITEM_UNIT_PRICE <> 0 GROUP BY p.PROJECT_ID, p.SUPPLIER_ID, p.FORM_NAME ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }
        //取得個別工資廠商合約資料與金額
        public List<plansummary> getPlanContract4Wage(string projectid)
        {
            List<plansummary> lst = new List<plansummary>();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT  p.PROJECT_ID + p.MAN_SUPPLIER_ID + p.MAN_FORM_NAME AS CONTRACT_ID, p.MAN_SUPPLIER_ID, p.MAN_FORM_NAME, " +
                    "SUM(p.ITEM_QUANTITY * ISNULL(p.MAN_PRICE,0)) WAGE_COST, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY p.MAN_SUPPLIER_ID) AS NO FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid and p.MAN_PRICE IS NOT NULL " +
                    "AND p.MAN_PRICE <> 0 GROUP BY p.PROJECT_ID, p.MAN_SUPPLIER_ID, p.MAN_FORM_NAME ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }
        //取得專案廠商合約之金額總計
        public plansummary getPlanContractAmount(string projectid)
        {
            plansummary lst = new plansummary();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT SUM(REVENUE) TOTAL_REVENUE, SUM(COST) TOTAL_COST, SUM(BUDGET) TOTAL_BUDGET, SUM(PROFIT) TOTAL_PROFIT " +
                    "FROM(select p.PROJECT_ID + p.SUPPLIER_ID + p.FORM_NAME AS CONTRACT_ID, " +
                    "p.SUPPLIER_ID, p.FORM_NAME, sum(p.ITEM_QUANTITY * p.ITEM_UNIT_COST) REVENUE, " +
                    "sum(p.ITEM_QUANTITY * p.ITEM_UNIT_COST * p.BUDGET_RATIO / 100) BUDGET, " +
                    "(sum(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) + SUM(p.ITEM_QUANTITY * ISNULL(p.MAN_PRICE,0))) COST, (sum(p.ITEM_QUANTITY * p.ITEM_UNIT_COST) - sum(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) - SUM(p.ITEM_QUANTITY * ISNULL(p.MAN_PRICE,0))) PROFIT, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY p.SUPPLIER_ID) AS NO FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid and p.ITEM_UNIT_PRICE IS NOT NULL " +
                    "AND p.ITEM_UNIT_PRICE <> 0 GROUP BY p.PROJECT_ID, p.SUPPLIER_ID, p.FORM_NAME)A ; "
                   , new SqlParameter("projectid", projectid)).First();
            }
            return lst;
        }

        //取得採購競標之詢價單資料
        public List<purchasesummary> getPurchaseForm4Offer(string projectid, string formname, string iswage)
        {

            logger.Info("search purchase form by 採購項目 =" + formname);
            List<purchasesummary> lstForm = new List<purchasesummary>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT C.code1 AS FORM_NAME, C.INQUIRY_FORM_ID as INQUIRY_FORM_ID, C.SUPPLIER_ID AS SUPPLIER_ID, D.TOTAL_ROWS AS TOTALROWS, D.TOTAL_PRICE AS TAmount " +
                         "FROM (select p.SUPPLIER_ID,  p.INQUIRY_FORM_ID, p.FORM_NAME AS code1, ISNULL(STATUS,'有效') STATUS, ISNULL(ISWAGE,'N')ISWAGE FROM PLAN_SUP_INQUIRY p LEFT OUTER JOIN PLAN_SUP_INQUIRY_ITEM pi " +
                         "ON p.INQUIRY_FORM_ID = pi.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL AND ISNULL(STATUS,'有效') <> '註銷' AND ISWAGE <> 'Y' GROUP BY p.FORM_NAME, p.INQUIRY_FORM_ID, p.STATUS, " +
                         "p.SUPPLIER_ID, p.ISWAGE HAVING p.FORM_NAME NOT IN (SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid " +
                         "AND p.ITEM_UNIT_PRICE IS NOT NULL AND p.ITEM_UNIT_PRICE <> 0 GROUP BY p.FORM_NAME))C LEFT OUTER JOIN " +
                         "(select  B.type, B.INQUIRY_FORM_ID, B.TOTAL_ROW AS TOTAL_ROWS, B.TOTALPRICE AS TOTAL_PRICE FROM (select p.FORM_NAME as type, p.INQUIRY_FORM_ID " +
                         "from PLAN_SUP_INQUIRY_ITEM pi LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL " +
                         "and pi.ITEM_UNIT_PRICE is not null GROUP BY p.INQUIRY_FORM_ID, p.FORM_NAME HAVING p.FORM_NAME NOT IN " +
                         "(SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.ITEM_UNIT_PRICE IS NOT NULL AND p.ITEM_UNIT_PRICE <> 0 GROUP BY p.FORM_NAME)) A " +
                         "RIGHT OUTER JOIN (select p.FORM_NAME as type, p.INQUIRY_FORM_ID, " +
                         "count(*) TOTAL_ROW, sum(ITEM_QTY * ITEM_UNIT_PRICE) TOTALPRICE from PLAN_SUP_INQUIRY_ITEM pi LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL GROUP BY p.INQUIRY_FORM_ID, " +
                         "p.FORM_NAME) B ON A.INQUIRY_FORM_ID + A.type = B.INQUIRY_FORM_ID + B.type) D ON C.INQUIRY_FORM_ID + C.code1 = D.INQUIRY_FORM_ID + D.type ";
            ;

            if (iswage == "Y")
            {
                sql = "SELECT C.code1 AS FORM_NAME, C.INQUIRY_FORM_ID as INQUIRY_FORM_ID, C.SUPPLIER_ID AS SUPPLIER_ID, D.TOTAL_ROWS AS TOTALROWS, D.TOTAL_PRICE AS TAmount " +
                         "FROM (select p.SUPPLIER_ID,  p.INQUIRY_FORM_ID, p.FORM_NAME AS code1, ISNULL(STATUS,'有效') STATUS, ISNULL(ISWAGE,'N')ISWAGE FROM PLAN_SUP_INQUIRY p LEFT OUTER JOIN PLAN_SUP_INQUIRY_ITEM pi " +
                         "ON p.INQUIRY_FORM_ID = pi.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL AND ISNULL(STATUS,'有效') <> '註銷' AND ISWAGE ='Y' GROUP BY p.FORM_NAME, p.INQUIRY_FORM_ID, p.STATUS, " +
                         "p.SUPPLIER_ID, p.ISWAGE HAVING p.FORM_NAME NOT IN (SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid " +
                         "AND p.MAN_PRICE IS NOT NULL AND p.MAN_PRICE <> 0 GROUP BY p.FORM_NAME))C LEFT OUTER JOIN " +
                         "(select  B.type, B.INQUIRY_FORM_ID, B.TOTAL_ROW AS TOTAL_ROWS, B.TOTALPRICE AS TOTAL_PRICE FROM (select p.FORM_NAME as type, p.INQUIRY_FORM_ID " +
                         "from PLAN_SUP_INQUIRY_ITEM pi LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL " +
                         "and pi.ITEM_UNIT_PRICE is not null GROUP BY p.INQUIRY_FORM_ID, p.FORM_NAME HAVING p.FORM_NAME NOT IN " +
                         "(SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.MAN_PRICE IS NOT NULL AND p.MAN_PRICE <> 0 GROUP BY p.FORM_NAME)) A " +
                         "RIGHT OUTER JOIN (select p.FORM_NAME as type, p.INQUIRY_FORM_ID, " +
                         "count(*) TOTAL_ROW, sum(ITEM_QTY * ITEM_UNIT_PRICE) TOTALPRICE from PLAN_SUP_INQUIRY_ITEM pi LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL GROUP BY p.INQUIRY_FORM_ID, " +
                         "p.FORM_NAME) B ON A.INQUIRY_FORM_ID + A.type = B.INQUIRY_FORM_ID + B.type) D ON C.INQUIRY_FORM_ID + C.code1 = D.INQUIRY_FORM_ID + D.type ";
                ;
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));

            //採購項目查詢條件
            if (null != formname && formname != "")
            {
                sql = sql + "WHERE C.code1 =@formname ";
                parameters.Add(new SqlParameter("formname", formname));
            }
            sql = sql + " ORDER BY C.code1;";
            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase form sql=" + sql);
                lstForm = context.Database.SqlQuery<purchasesummary>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase form count=" + lstForm.Count);
            return lstForm;
        }

        //取得特定專案報價之供應商資料
        public List<COMPARASION_DATA_4PLAN> getComparisonData(string projectid, string typecode1, string typecode2, string systemMain, string systemSub, string formName, string iswage)
        {
            List<COMPARASION_DATA_4PLAN> lst = new List<COMPARASION_DATA_4PLAN>();
            string sql = "SELECT  pfItem.INQUIRY_FORM_ID AS INQUIRY_FORM_ID, " +
                "f.SUPPLIER_ID as SUPPLIER_NAME, f.FORM_NAME AS FORM_NAME, ISNULL(STATUS,'有效') STATUS, SUM(pfitem.ITEM_UNIT_PRICE*pfitem.ITEM_QTY) as TAmount " +
                "FROM PLAN_ITEM pItem LEFT OUTER JOIN " +
                "PLAN_SUP_INQUIRY_ITEM pfItem ON pItem.PLAN_ITEM_ID = pfItem.PLAN_ITEM_ID " +
                "inner join PLAN_SUP_INQUIRY f on pfItem.INQUIRY_FORM_ID = f.INQUIRY_FORM_ID " +
                "WHERE pItem.PROJECT_ID = @projectid AND f.SUPPLIER_ID is not null AND ISNULL(STATUS,'有效') <> '註銷' AND ISNULL(f.ISWAGE,'N')=@iswage  ";
            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            //設定報價單條件，材料或工資
            parameters.Add(new SqlParameter("iswage", iswage));
            //九宮格條件
            if (null != typecode1 && "" != typecode1)
            {
                //sql = sql + " AND pItem.TYPE_CODE_1='" + typecode1 + "'";
                sql = sql + " AND pItem.TYPE_CODE_1=@typecode1";
                parameters.Add(new SqlParameter("typecode1", typecode1));
            }
            //次九宮格條件
            if (null != typecode2 && "" != typecode2)
            {
                //sql = sql + " AND pItem.TYPE_CODE_2='" + typecode2 + "'";
                sql = sql + " AND pItem.TYPE_CODE_2=@typecode2";
                parameters.Add(new SqlParameter("typecode2", typecode2));
            }
            //主系統條件
            if (null != systemMain && "" != systemMain)
            {
                // sql = sql + " AND pItem.SYSTEM_MAIN='" + systemMain + "'";
                sql = sql + " AND pItem.SYSTEM_MAIN=@systemMain";
                parameters.Add(new SqlParameter("systemMain", systemMain));
            }
            //次系統條件
            if (null != systemSub && "" != systemSub)
            {
                //sql = sql + " AND pItem.SYSTEM_SUB='" + systemSub + "'";
                sql = sql + " AND pItem.SYSTEM_SUB=@systemSub";
                parameters.Add(new SqlParameter("systemSub", systemSub));
            }
            //採購名稱條件
            if (null != formName && "" != formName)
            {
                //sql = sql + " AND f.FORM_NAME='" + formName + "'";
                sql = sql + " AND f.FORM_NAME = @formName ";
                parameters.Add(new SqlParameter("formName", formName));
            }
            sql = sql + " GROUP BY pfItem.INQUIRY_FORM_ID ,f.SUPPLIER_ID, f.FORM_NAME, f.STATUS ;";
            logger.Info("comparison data sql=" + sql);
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<COMPARASION_DATA_4PLAN>(sql, parameters.ToArray()).ToList();
                logger.Info("Get ComparisonData Count=" + lst.Count);
            }
            return lst;
        }

        //比價資料
        public DataTable getComparisonDataToPivot(string projectid, string typecode1, string typecode2, string systemMain, string systemSub, string formName, string iswage)
        {
            //採購名稱條件
            if (null != formName && "" != formName)
            {
                string sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位," +
                "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as SUPPLIER_NAME, " +
                "pitem.ITEM_UNIT_PRICE 單價, " +
                "(SELECT FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as FORM_NAME, fitem.ITEM_UNIT_PRICE  " +
                "from PLAN_ITEM pitem " +
                "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                "where pitem.PROJECT_ID = @projectid ";

                if (iswage == "Y")
                {
                    sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位," +
                    "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID) as SUPPLIER_NAME, " +
                    "pitem.MAN_PRICE 工資單價, " +
                    "(SELECT FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as FORM_NAME, fitem.ITEM_UNIT_PRICE  " +
                    "from PLAN_ITEM pitem " +
                    "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                    " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                    "where pitem.PROJECT_ID = @projectid ";
                }

                var parameters = new Dictionary<string, Object>();
                //設定專案名編號資料
                parameters.Add("projectid", projectid);
                //九宮格條件
                if (null != typecode1 && "" != typecode1)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_1='" + typecode1 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_1=@typecode1";
                    parameters.Add("typecode1", typecode1);
                }
                //次九宮格條件
                if (null != typecode2 && "" != typecode2)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_2='" + typecode2 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_2=@typecode2";
                    parameters.Add("typecode2", typecode2);
                }
                //主系統條件
                if (null != systemMain && "" != systemMain)
                {
                    // sql = sql + " AND pItem.SYSTEM_MAIN='" + systemMain + "'";
                    sql = sql + " AND pItem.SYSTEM_MAIN=@systemMain";
                    parameters.Add("systemMain", systemMain);
                }
                //次系統條件
                if (null != systemSub && "" != systemSub)
                {
                    //sql = sql + " AND pItem.SYSTEM_SUB='" + systemSub + "'";
                    sql = sql + " AND pItem.SYSTEM_SUB=@systemSub";
                    parameters.Add("systemSub", systemSub);
                }

                //取的欄位維度條件
                List<COMPARASION_DATA_4PLAN> lstSuppluerQuo = getComparisonData(projectid, typecode1, typecode2, systemMain, systemSub, formName, iswage);
                if (lstSuppluerQuo.Count == 0)
                {
                    throw new Exception("相關條件沒有任何報價資料!!");
                }
                //設定供應商報價資料，供前端畫面調用
                dirSupplierQuo = new Dictionary<string, COMPARASION_DATA_4PLAN>();
                string dimString = "";
                foreach (var it in lstSuppluerQuo)
                {
                    logger.Debug("Supplier=" + it.SUPPLIER_NAME + "," + it.INQUIRY_FORM_ID + "," + it.FORM_NAME);
                    if (dimString == "")
                    {
                        dimString = "[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    else
                    {
                        dimString = dimString + ",[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    //設定供應商報價資料，供前端畫面調用
                    dirSupplierQuo.Add(it.INQUIRY_FORM_ID, it);
                }

                logger.Debug("dimString=" + dimString);
                //sql = sql + " AND FORM_NAME ='" + formName + "'";
                sql = sql + ") souce pivot(MIN(ITEM_UNIT_PRICE) FOR SUPPLIER_NAME IN(" + dimString + ")) as pvt WHERE FORM_NAME =@formName ORDER BY 行數; ";
                parameters.Add("formName", formName);
                logger.Info("comparison data sql=" + sql);
                DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
                //Pivot pvt = new Pivot(ds.Tables[0]);
                return ds.Tables[0];
            }
            else
            {
                string sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位," +
                "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as SUPPLIER_NAME, " +
                "pitem.ITEM_UNIT_PRICE 單價, fitem.ITEM_UNIT_PRICE " +
                "from PLAN_ITEM pitem " +
                "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                "where pitem.PROJECT_ID = @projectid ";

                if (iswage == "Y")
                {
                    sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位," +
                    "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID) as SUPPLIER_NAME, " +
                    "pitem.MAN_PRICE 工資單價, fitem.ITEM_UNIT_PRICE " +
                    "from PLAN_ITEM pitem " +
                    "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                    " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                    "where pitem.PROJECT_ID = @projectid ";
                }
                var parameters = new Dictionary<string, Object>();
                //設定專案名編號資料
                parameters.Add("projectid", projectid);
                //九宮格條件
                if (null != typecode1 && "" != typecode1)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_1='" + typecode1 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_1=@typecode1";
                    parameters.Add("typecode1", typecode1);
                }
                //次九宮格條件
                if (null != typecode2 && "" != typecode2)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_2='" + typecode2 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_2=@typecode2";
                    parameters.Add("typecode2", typecode2);
                }
                //主系統條件
                if (null != systemMain && "" != systemMain)
                {
                    // sql = sql + " AND pItem.SYSTEM_MAIN='" + systemMain + "'";
                    sql = sql + " AND pItem.SYSTEM_MAIN=@systemMain";
                    parameters.Add("systemMain", systemMain);
                }
                //次系統條件
                if (null != systemSub && "" != systemSub)
                {
                    //sql = sql + " AND pItem.SYSTEM_SUB='" + systemSub + "'";
                    sql = sql + " AND pItem.SYSTEM_SUB=@systemSub";
                    parameters.Add("systemSub", systemSub);
                }

                //取的欄位維度條件
                List<COMPARASION_DATA_4PLAN> lstSuppluerQuo = getComparisonData(projectid, typecode1, typecode2, systemMain, systemSub, formName, iswage);
                if (lstSuppluerQuo.Count == 0)
                {
                    throw new Exception("相關條件沒有任何報價資料!!");
                }
                //設定供應商報價資料，供前端畫面調用
                dirSupplierQuo = new Dictionary<string, COMPARASION_DATA_4PLAN>();
                string dimString = "";
                foreach (var it in lstSuppluerQuo)
                {
                    logger.Debug("Supplier=" + it.SUPPLIER_NAME + "," + it.INQUIRY_FORM_ID + "," + it.FORM_NAME);
                    if (dimString == "")
                    {
                        dimString = "[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    else
                    {
                        dimString = dimString + ",[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    //設定供應商報價資料，供前端畫面調用
                    dirSupplierQuo.Add(it.INQUIRY_FORM_ID, it);
                }

                logger.Debug("dimString=" + dimString);
                //sql = sql + " AND FORM_NAME ='" + formName + "'";
                sql = sql + ") souce pivot(MIN(ITEM_UNIT_PRICE) FOR SUPPLIER_NAME IN(" + dimString + ")) as pvt ORDER BY 行數; ";
                logger.Info("comparison data sql=" + sql);
                DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
                //Pivot pvt = new Pivot(ds.Tables[0]);
                return ds.Tables[0];
            }
        }

        public int addSuplplierFormFromQuote(string formid)
        {
            int i = 0;
            logger.Info("add plan supplier form from Quote by form id" + formid);
            string sql = "UPDATE  PLAN_SUP_INQUIRY SET COUNTER_OFFER = 'Y', MODIFY_DATE = getdate()  WHERE INQUIRY_FORM_ID=@formid ";
            logger.Debug("add form from Quote sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        public int removeSuplplierFormFromQuote(string formid)
        {
            int i = 0;
            logger.Info("Remove plan supplier form from Quote by form id" + formid);
            string sql = "UPDATE  PLAN_SUP_INQUIRY SET STATUS = '註銷', MODIFY_DATE = getdate() WHERE INQUIRY_FORM_ID=@formid ";
            logger.Debug("remove form from Quote sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }

        public int filterSuplplierFormFromQuote(string formid)
        {
            int i = 0;
            logger.Info("Filter plan supplier form from Quote by form id" + formid);
            string sql = "UPDATE  PLAN_SUP_INQUIRY SET COUNTER_OFFER = 'M', MODIFY_DATE = getdate() WHERE INQUIRY_FORM_ID=@formid ";
            logger.Debug("remove form from Quote sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //由報價單資料更新標單資料
        public int updateCostFromQuote(string planItemid, decimal price, string iswage)
        {
            int i = 0;
            logger.Info("Update Cost:plan item id=" + planItemid + ",price=" + price);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_ITEM SET ITEM_UNIT_PRICE =@price WHERE PLAN_ITEM_ID=@pitemid ";
            //將工資報價單更新工資報價欄位
            if (iswage == "Y")
            {
                sql = "UPDATE PLAN_ITEM SET MAN_PRICE=@price WHERE PLAN_ITEM_ID=@pitemid ";
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("price", price));
            parameters.Add(new SqlParameter("pitemid", planItemid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Info("Update Cost:" + i);
            return i;
        }
        public int batchUpdateCostFromQuote(string formid, string iswage)
        {
            int i = 0;
            logger.Info("Copy cost from Quote to Tnd by form id" + formid);
            string sql = "UPDATE  PLAN_ITEM SET item_unit_price = i.ITEM_UNIT_PRICE, supplier_id = i.SUPPLIER_ID, form_name = i.FORM_NAME " +
                "FROM(select i.plan_item_id, fi.ITEM_UNIT_PRICE, fi.INQUIRY_FORM_ID, pf.SUPPLIER_ID, pf.FORM_NAME from PLAN_ITEM i " +
                ", PLAN_SUP_INQUIRY_ITEM fi, PLAN_SUP_INQUIRY pf " +
               "where i.PLAN_ITEM_ID = fi.PLAN_ITEM_ID and fi.INQUIRY_FORM_ID = pf.INQUIRY_FORM_ID and fi.INQUIRY_FORM_ID = @formid) i " +
                "WHERE  i.plan_item_id = PLAN_ITEM.PLAN_ITEM_ID ";

            //將工資報價單更新工資報價欄位
            if (iswage == "Y")
            {
                sql = "UPDATE  PLAN_ITEM SET man_price = i.ITEM_UNIT_PRICE, man_supplier_id = i.SUPPLIER_ID, man_form_name = i.FORM_NAME "
                + "FROM (select i.plan_item_id, fi.ITEM_UNIT_PRICE, fi.INQUIRY_FORM_ID, pf.SUPPLIER_ID, pf.FORM_NAME from PLAN_ITEM i "
                + ", PLAN_SUP_INQUIRY_ITEM fi, PLAN_SUP_INQUIRY pf " 
                + "where i.PLAN_ITEM_ID = fi.PLAN_ITEM_ID and fi.INQUIRY_FORM_ID = pf.INQUIRY_FORM_ID and fi.INQUIRY_FORM_ID = @formid) i " 
                + "WHERE  i.plan_item_id = PLAN_ITEM.PLAN_ITEM_ID ";
            }
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }

        public int addContractId(string projectid)
        {
            int i = 0;
            //將合約編號寫入PLAN PAYMENT TERMS
            logger.Info("copy contract id into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                List<topmeperp.Models.PLAN_PAYMENT_TERMS> lstItem = new List<PLAN_PAYMENT_TERMS>();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID) " +
                       "SELECT distinct ('" + projectid + "' + p.SUPPLIER_ID + p.FORM_NAME) AS contractid, '" + projectid + "' FROM  PLAN_ITEM p WHERE p.SUPPLIER_ID IS NOT NULL " +
                       "AND '" + projectid + "' + p.SUPPLIER_ID + p.FORM_NAME NOT IN(SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";

                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }

        public List<PLAN_ITEM> getContractItemsByContractName(string contractid)
        {
            List<PLAN_ITEM> lst = new List<PLAN_ITEM>();
            using (var context = new topmepEntities())
            {
                lst = context.PLAN_ITEM.SqlQuery("SELECT * FROM PLAN_ITEM WHERE  PROJECT_ID + SUPPLIER_ID + FORM_NAME = @contractid ;"
                    , new SqlParameter("contractid", contractid)).ToList();
            }
            logger.Info("get plan supplier contract items count:" + lst.Count);
            return lst;
        }

        public PLAN_PAYMENT_TERMS getPaymentTerm(string contractid)
        {
            logger.Debug("get payment terms by contractid=" + contractid);
            PLAN_PAYMENT_TERMS payment = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                payment = context.PLAN_PAYMENT_TERMS.SqlQuery("SELECT * FROM PLAN_PAYMENT_TERMS WHERE CONTRACT_ID=@contractid",
                new SqlParameter("contractid", contractid)).FirstOrDefault();
            }
            return payment;
        }

        public List<PLAN_ITEM> getPendingItems(string projectid)
        {
            List<PLAN_ITEM> lst = new List<PLAN_ITEM>();
            using (var context = new topmepEntities())
            {
                // ITEM_UNIT IS NOT NULL(確認單位欄位是空值就是不需採購的欄位嗎) AND ITEM_UNIT_PRICE IS NULL
                lst = context.Database.SqlQuery<PLAN_ITEM>("SELECT * FROM PLAN_ITEM WHERE PROJECT_ID =@projectid AND ITEM_UNIT IS NOT NULL AND ITEM_UNIT_PRICE IS NULL  ;"
                    , new SqlParameter("projectid", projectid)).ToList();
            }
            logger.Info("get plan pending items count:" + lst.Count);
            return lst;
        }

        public int updatePaymentTerms(PLAN_PAYMENT_TERMS item)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_PAYMENT_TERMS.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("updatePaymentTerms fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        public int changePlanFormStatus(string formid, string status)
        {
            int i = 0;
            logger.Info("Update plan sup inquiry form status formid=" + formid + ",status=" + status);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_SUP_INQUIRY SET STATUS=@status WHERE INQUIRY_FORM_ID=@formid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", status));
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Debug("Update plan sup inquiry form status  :" + i);
            return i;
        }
        //更新得標標單品項之採購前置天數
        public int updateLeadTime(string projectid, List<PLAN_ITEM> lstItem)
        {
            //1.新增前置天數資料
            int i = 0;
            logger.Info("update lead time = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_ITEM item in lstItem)
                {
                    PLAN_ITEM existItem = null;
                    logger.Debug("plan item id=" + item.PLAN_ITEM_ID);
                    if (item.PLAN_ITEM_ID != null && item.PLAN_ITEM_ID != "")
                    {
                        existItem = context.PLAN_ITEM.Find(item.PLAN_ITEM_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("projectid", projectid));
                        parameters.Add(new SqlParameter("planitemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ITEM WHERE PROJECT_ID = @projectid and PLAN_ITEM_ID = @planitemid ";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.PLAN_ITEM_ID);
                        PLAN_ITEM excelItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ITEM.Find(excelItem.PLAN_ITEM_ID);

                    }
                    logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                    existItem.LEAD_TIME = item.LEAD_TIME;
                    existItem.MODIFY_USER_ID = item.MODIFY_USER_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_ITEM.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update lead time count =" + i);
            return i;
        }
    }
}
