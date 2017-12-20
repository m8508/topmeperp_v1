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
using System.Globalization;


namespace topmeperp.Service
{
    public class PlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string message = "";
        public TND_PROJECT project = null;
        public TND_PROJECT budgetTable = null;


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

                        string sql = "SELECT * FROM PLAN_BUDGET WHERE PROJECT_ID = @projectid and ISNULL(TYPE_CODE_1, '') + ISNULL(TYPE_CODE_2, '') = @code1 + @code2";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.TYPE_CODE_1 + item.TYPE_CODE_2 + item.SYSTEM_MAIN + item.SYSTEM_SUB);
                        PLAN_BUDGET excelItem = context.PLAN_BUDGET.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_BUDGET.Find(excelItem.PLAN_BUDGET_ID);

                    }
                    logger.Debug("find exist item=" + existItem.PLAN_BUDGET_ID);
                    existItem.BUDGET_RATIO = item.BUDGET_RATIO;
                    existItem.BUDGET_WAGE_RATIO = item.BUDGET_WAGE_RATIO;
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
            string sql = "UPDATE PLAN_ITEM SET PLAN_ITEM.BUDGET_RATIO = plan_budget.BUDGET_RATIO, PLAN_ITEM.BUDGET_WAGE_RATIO = plan_budget.BUDGET_WAGE_RATIO " +
                   "from PLAN_ITEM inner join " +
                   "plan_budget on REPLACE(PLAN_ITEM.PROJECT_ID, ' ', '') + IIF(REPLACE(PLAN_ITEM.TYPE_CODE_1, ' ', '') is null, '', IIF(REPLACE(PLAN_ITEM.TYPE_CODE_1, ' ', '') = 0, '', REPLACE(PLAN_ITEM.TYPE_CODE_1, ' ', ''))) + IIF(REPLACE(PLAN_ITEM.TYPE_CODE_2, ' ', '') is null, '', IIF(REPLACE(PLAN_ITEM.TYPE_CODE_2, ' ', '') = 0, '', REPLACE(PLAN_ITEM.TYPE_CODE_2, ' ', ''))) " +
                   "= REPLACE(plan_budget.PROJECT_ID, ' ', '') + IIF(REPLACE(plan_budget.TYPE_CODE_1, ' ', '') is null, '', IIF(REPLACE(plan_budget.TYPE_CODE_1, ' ', '') = 0, '', REPLACE(plan_budget.TYPE_CODE_1, ' ', ''))) + IIF(REPLACE(plan_budget.TYPE_CODE_2, ' ', '') is null, '', IIF(REPLACE(plan_budget.TYPE_CODE_2, ' ', '') = 0, '', REPLACE(plan_budget.TYPE_CODE_2, ' ', ''))) WHERE PLAN_ITEM.PROJECT_ID  = @id ";
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
                plan = context.Database.SqlQuery<PlanRevenue>("SELECT p.PROJECT_ID AS CONTRACT_ID, " +
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
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID, TYPE) " +
                       "SELECT '" + projectid + "' AS contractid, '" + projectid + "', 'O' FROM TND_PROJECT p WHERE p.PROJECT_ID = '" + projectid + "'  " +
                       "AND '" + projectid + "' NOT IN(SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";
                logger.Info("sql =" + sql);
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
                string sql = "SELECT C.*, SUM(ISNULL(MATERIAL_COST_INMAP,0) * BUDGET / 100) + SUM(ISNULL(MAN_DAY_INMAP,0) * BUDGET_WAGE / 100) AS AMOUNT_BY_CODE FROM " +
                    "(SELECT MAINCODE, MAINCODE_DESC, SUB_CODE, SUB_DESC, MATERIAL_COST_INMAP, MAN_DAY_INMAP, CONTRACT_PRICE, "
                    + "ISNULL(BUDGET_RATIO, 0) as BUDGET, ISNULL(BUDGET_WAGE_RATIO, 0) as BUDGET_WAGE, COST_RATIO FROM (SELECT" +
                    "(select TYPE_CODE_1 + TYPE_CODE_2 from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE, " +
                    "(select TYPE_DESC from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE_DESC ," +
                    "(select SUB_TYPE_ID from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) T_SUB_CODE, " +
                    "TYPE_CODE_2 SUB_CODE," +
                    "(select TYPE_DESC from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) SUB_DESC, " +
                    "SUM(MapQty * tndPrice) MATERIAL_COST_INMAP, SUM(MapQty * RATIO * WagePrice) MAN_DAY_INMAP, SUM(ITEM_UNIT_PRICE * tndPrice) CONTRACT_PRICE, count(*) ITEM_COUNT " +
                    "FROM (SELECT pi.*, w.RATIO, w.PRICE, map.QTY MapQty, ISNULL(p.WAGE_MULTIPLIER, 0) AS WagePrice, it.ITEM_UNIT_PRICE AS tndPrice, it.ITEM_QUANTITY AS tndQTY FROM PLAN_ITEM pi LEFT OUTER JOIN TND_WAGE w " +
                    "ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT OUTER JOIN TND_PROJECT_ITEM it ON it.PROJECT_ITEM_ID = pi.PLAN_ITEM_ID " +
                    "LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID WHERE it.project_id = @projectid) A " +
                    "GROUP BY TYPE_CODE_1, TYPE_CODE_2) B LEFT OUTER JOIN (SELECT p.TYPE_CODE_1, p.TYPE_CODE_2, SUM(p.BUDGET_RATIO*map.QTY)/SUM(map.QTY) BUDGET_RATIO, " +
                    "SUM(p.BUDGET_WAGE_RATIO*map.QTY)/SUM(map.QTY) BUDGET_WAGE_RATIO, " +
                    "SUM(p.TND_RATIO*map.QTY)/SUM(map.QTY) COST_RATIO FROM PLAN_ITEM p LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON p.PLAN_ITEM_ID = map.PROJECT_ITEM_ID WHERE p.PROJECT_ID =@projectid GROUP BY p.TYPE_CODE_1, p.TYPE_CODE_2) D " +
                    "ON IIF(MAINCODE is null, '', IIF(MAINCODE = 0, '', MAINCODE)) + IIF(SUB_CODE is null, '', IIF(SUB_CODE = 0, '', SUB_CODE)) = IIF(REPLACE(D.TYPE_CODE_1, ' ', '') is null, '', IIF(REPLACE(D.TYPE_CODE_1, ' ', '') = 0, '', REPLACE(D.TYPE_CODE_1, ' ', ''))) + IIF(REPLACE(D.TYPE_CODE_2, ' ', '') is null, '', IIF(REPLACE(D.TYPE_CODE_2, ' ', '') = 0, '', REPLACE(D.TYPE_CODE_2, ' ', ''))) " +
                    ") C GROUP BY MAINCODE, MAINCODE_DESC, SUB_CODE, SUB_DESC, MATERIAL_COST_INMAP, MAN_DAY_INMAP, CONTRACT_PRICE, BUDGET, BUDGET_WAGE, COST_RATIO ORDER BY ISNULL(MAINCODE, '無'), ISNULL(SUB_CODE, '無') ";
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
                string sql = "SELECT SUM(ISNULL(TND_COST,0)) AS TOTAL_COST, SUM(ISNULL(BUDGET,0)) AS MATERIAL_BUDGET, SUM(ISNULL(BUDGET_WAGE,0)) AS WAGE_BUDGET, SUM(ISNULL(BUDGET,0)) + SUM(ISNULL(BUDGET_WAGE,0)) AS TOTAL_BUDGET, SUM(ISNULL(P_COST,0)) AS TOTAL_P_COST FROM (SELECT(select TYPE_CODE_1 + TYPE_CODE_2 from REF_TYPE_MAIN WHERE  " +
                    "TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE, (select TYPE_DESC from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE_DESC, " +
                    "(select SUB_TYPE_ID from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) T_SUB_CODE, TYPE_CODE_2 SUB_CODE, " +
                    "(select TYPE_DESC from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) SUB_DESC, SUM(MapQty * tndPrice) MATERIAL_COST_INMAP, " +
                    "SUM(MapQty * RATIO * WagePrice) MAN_DAY_INMAP,count(*) ITEM_COUNT, SUM(MapQty * tndPrice * ISNULL(BUDGET_RATIO, 100)/100) BUDGET, SUM(MapQty * RATIO * WagePrice * ISNULL(BUDGET_WAGE_RATIO, 100)/100) BUDGET_WAGE, " +
                    "SUM(MapQty * ITEM_UNIT_COST) + SUM(MapQty * MAN_PRICE) P_COST, SUM(tndQTY * ITEM_UNIT_PRICE) TND_COST FROM " +
                    "(SELECT pi.*, w.RATIO, w.PRICE, map.QTY MapQty, ISNULL(p.WAGE_MULTIPLIER, 0) AS WagePrice, it.ITEM_UNIT_PRICE AS tndPrice, it.ITEM_QUANTITY AS tndQTY FROM PLAN_ITEM pi LEFT OUTER JOIN " +
                    "TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID LEFT OUTER JOIN TND_PROJECT_ITEM it " +
                    "ON it.PROJECT_ITEM_ID = pi.PLAN_ITEM_ID WHERE it.project_id = @projectid) A  " +
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
            string sql = "SELECT SUM(map.QTY*pi.ITEM_UNIT_PRICE) AS ITEM_COST, SUM(map.QTY*tpi.ITEM_UNIT_PRICE*ISNULL(pi.BUDGET_RATIO, 100)/100) AS ITEM_BUDGET, SUM(map.QTY*w.RATIO*ISNULL(p.WAGE_MULTIPLIER, 0)*ISNULL(pi.BUDGET_WAGE_RATIO, 100)/100) AS ITEM_BUDGET_WAGE FROM PLAN_ITEM pi " +
                "LEFT JOIN TND_PROJECT_ITEM tpi ON pi.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID LEFT OUTER JOIN TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID " +
                "LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID ";
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
        public PurchaseFormModel POFormData = null;
        public List<FIN_SUBJECT> ExpBudgetItem = null;

        #region 取得得標標單項目內容
        //取得標單品項資料
        public List<PlanItem4Map> getPlanItem(string checkEx, string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub, string formName, string supplier, string delFlg)
        {

            logger.Info("search plan item by 九宮格 =" + typeCode1 + "search plan item by 次九宮格 =" + typeCode2 + "search plan item by 主系統 =" + systemMain + "search plan item by 次系統 =" + systemSub + "search plan item by 採購項目 =" + formName + "search plan item by 材料供應商 =" + supplier);
            List<topmeperp.Models.PlanItem4Map> lstItem = new List<PlanItem4Map>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pi.*, map.QTY AS MAP_QTY FROM PLAN_ITEM pi LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID ";
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
            //顯示未分類資料
            if (null != checkEx && checkEx != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 is null or pi.TYPE_CODE_1='' ";
            }
            //刪除註記
            if ("*" != delFlg)
            {
                sql = sql + "AND ISNULL(pi.DEL_FLAG,'N')=@delFlg ";
                parameters.Add(new SqlParameter("delFlg", delFlg));
            }
            sql = sql + "  ORDER BY EXCEL_ROW_ID;";
            using (var context = new topmepEntities())
            {
                logger.Debug("get plan item sql=" + sql);
                lstItem = context.Database.SqlQuery<PlanItem4Map>(sql, parameters.ToArray()).ToList();
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
            if (null == item.PLAN_ITEM_ID || item.PLAN_ITEM_ID == "")
            {
                logger.Debug("add new plan item in porjectid=" + item.PROJECT_ID);
                item = getNewPlanItemID(item);
            }
            logger.Debug("plan item key=" + item.PLAN_ITEM_ID);
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

        private PLAN_ITEM getNewPlanItemID(PLAN_ITEM item)
        {
            string sql = "SELECT MAX(CAST(SUBSTRING(PLAN_ITEM_ID,8,LEN(PLAN_ITEM_ID)) AS INT) +1) MaxSN, MAX(EXCEL_ROW_ID) + 1 as Row "
                + " FROM PLAN_ITEM WHERE PROJECT_ID = @projectid ; ";
            var parameters = new Dictionary<string, Object>();
            parameters.Add("projectid", item.PROJECT_ID);
            DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
            logger.Debug("sql=" + sql + "," + ds.Tables[0].Rows[0][0].ToString() + "," + ds.Tables[0].Rows[0][1].ToString());
            int longMaxExcel = 1;
            int longMaxItem = 1;
            if (DBNull.Value != ds.Tables[0].Rows[0][0])
            {
                longMaxItem = int.Parse(ds.Tables[0].Rows[0][0].ToString());
                longMaxExcel = int.Parse(ds.Tables[0].Rows[0][1].ToString());
            }
            logger.Debug("new plan item id=" + longMaxItem + ",ExcelRowID=" + longMaxExcel);
            item.PLAN_ITEM_ID = item.PROJECT_ID + "-" + longMaxItem;
            //新品項不會有Excel Row_id
            if (null == item.EXCEL_ROW_ID || item.EXCEL_ROW_ID == 0)
            {
                item.EXCEL_ROW_ID = longMaxExcel;
            }
            return item;
        }

        //於現有品項下方新增一筆資料
        public int addPlanItemAfter(PLAN_ITEM item)
        {
            string sql = "UPDATE PLAN_ITEM SET EXCEL_ROW_ID=EXCEL_ROW_ID+1 WHERE PROJECT_ID = @projectid AND EXCEL_ROW_ID> @ExcelRowId ";

            using (var db = new topmepEntities())
            {
                logger.Debug("add exce rowid sql=" + sql + ",projectid=" + item.PROJECT_ID + ",ExcelRowI=" + item.EXCEL_ROW_ID);
                db.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", item.PROJECT_ID), new SqlParameter("ExcelRowId", item.EXCEL_ROW_ID));
            }
            item.PLAN_ITEM_ID = "";
            item.ITEM_UNIT_COST = null;
            item.EXCEL_ROW_ID = item.EXCEL_ROW_ID + 1;
            return updatePlanItem(item);
        }
        //將Plan Item 註記刪除
        public int changePlanItem(string itemid, string delFlag)
        {
            string sql = "UPDATE PLAN_ITEM SET DEL_FLAG=@delFlag WHERE PLAN_ITEM_ID = @itemid";
            int i = 0;
            using (var db = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("itemid", itemid));
                parameters.Add(new SqlParameter("delFlag", delFlag));
                logger.Info("Update PLAN_ITEM FLAG=" + sql + ",itemid=" + itemid + ",delFlag=" + delFlag);
                i = db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
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
                //0.清除所有空白詢價單樣板//僅刪除材料之空白詢價單
                string sql = "DELETE FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID IN (SELECT INQUIRY_FORM_ID FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID=@projectid AND ISNULL(ISWAGE,'N')='N');";
                i2 = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", projectid));
                logger.Info("delete template inquiry form item  by porjectid=" + projectid + ",result=" + i2);
                sql = "DELETE FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID=@projectid AND ISNULL(ISWAGE,'N')='N'; ";
                i2 = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", projectid));
                logger.Info("delete template inquiry form  by porjectid=" + projectid + ",result=" + i2);

                //1.依據專案取得九宮格次九宮格分類.
                sql = "SELECT DISTINCT isnull(TYPE_CODE_1,'未分類') TYPE_CODE_1," +
                   "(SELECT TYPE_DESC FROM REF_TYPE_MAIN m WHERE m.TYPE_CODE_1 + m.TYPE_CODE_2 = p.TYPE_CODE_1) as TYPE_CODE_1_NAME, " +
                   "isnull(TYPE_CODE_2,'未分類') TYPE_CODE_2," +
                   "(SELECT TYPE_DESC FROM REF_TYPE_SUB sub WHERE sub.TYPE_CODE_ID = p.TYPE_CODE_1 AND sub.SUB_TYPE_CODE = p.TYPE_CODE_2) as TYPE_CODE_2_NAME " +
                   "FROM PLAN_ITEM p WHERE PROJECT_ID = @projectid ORDER BY TYPE_CODE_1 ,Type_CODE_2; ";

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
                    + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QTY, ITEM_UNIT_PRICE, ITEM_REMARK,ITEM_ID) "
                    + "SELECT '" + form.INQUIRY_FORM_ID + "' as INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE_1 AS TYPE_CODE, "
                    + "TYPE_CODE_2 AS SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, map.QTY, ITEM_UNIT_COST, ITEM_REMARK,pi.ITEM_ID ITEM_ID "
                    + "FROM PLAN_ITEM pi LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID where PLAN_ITEM_ID IN (" + ItemId + ")";
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
                formInquiryItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery("SELECT i.[INQUIRY_ITEM_ID],i.[INQUIRY_FORM_ID]" +
                    ", i.[PLAN_ITEM_ID], i.[TYPE_CODE], i.[SUB_TYPE_CODE], pi.[ITEM_ID], i.[ITEM_DESC], i.[ITEM_UNIT] "
                    + " , i.[ITEM_QTY],i.[ITEM_UNIT_PRICE], i.[ITEM_QTY_ORG] , i.[ITEM_UNITPRICE_ORG], i.ITEM_REMARK "
                    + " , i.[MODIFY_ID], i.[MODIFY_DATE], i.[WAGE_PRICE]  "
                    + "FROM PLAN_SUP_INQUIRY_ITEM i LEFT OUTER JOIN  PLAN_ITEM pi on i.PLAN_ITEM_ID = pi.PLAN_ITEM_ID "
                    + "WHERE i.INQUIRY_FORM_ID=@formid ORDER BY pi.EXCEL_ROW_ID", new SqlParameter("formid", formid)).ToList();
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
                ZipFileCreator.CreateDirectory(tempFolder + formInquiry.FORM_NAME);
                string fileLocation = poi.exportExcel4po(formInquiry, formInquiryItem, true, false);
                logger.Debug("temp file=" + fileLocation);
            }
            //4.Zip all file
            return zipTool.ZipDirectory(tempFolder);
            //return zipTool.ZipFiles(tempFolder, null, p.PROJECT_NAME);
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
        //採發階段發包分項與預算 (材料預算)
        public PurchaseFormModel getInquiryWithBudget(TND_PROJECT project, string status)
        {
            POFormData = new PurchaseFormModel();
            getBudgetSummary(project);
            POFormData.materialTemplateWithBudget = getTemplateRefBudget(project, "N", status);
            POFormData.wageTemplateWithBudget = getTemplateRefBudget(project, "Y", status);
            return POFormData;
        }
        //取得詢價單樣本與分項預算
        public IEnumerable<PURCHASE_ORDER> getTemplateRefBudget(TND_PROJECT project, string iswage, string status)
        {
            logger.Info("get purchase template by projectid=" + project.PROJECT_ID);
            List<PURCHASE_ORDER> lst = new List<PURCHASE_ORDER>();
            string sql = "";
            decimal wageunitprice = 2500;
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", project.PROJECT_ID));
                parameters.Add(new SqlParameter("status", status));
                if (iswage == "N")
                {
                    //取得詢價單樣本資訊 - 材料預算-圖算數量*報價標單(Project_item)單價 * 預算折扣比率
                    sql = "SELECT tmp.*,CountPO, (SELECT SUM(v.QTY * tpi.ITEM_UNIT_PRICE * it.BUDGET_RATIO/100 ) as BudgetAmount FROM PLAN_ITEM it LEFT JOIN vw_MAP_MATERLIALIST v ON it.PLAN_ITEM_ID = v.PROJECT_ITEM_ID  "
                       + "LEFT JOIN TND_PROJECT_ITEM tpi ON it.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID "
                       + "WHERE it.PLAN_ITEM_ID in (SELECT  iit.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY_ITEM iit WHERE iit.INQUIRY_FORM_ID = tmp.INQUIRY_FORM_ID)) AS BudgetAmount "
                       + "FROM(SELECT * FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID is Null AND PROJECT_ID = @projectid AND ISNULL(STATUS,'有效')=@status AND ISNULL(ISWAGE,'N')='N') tmp LEFT OUTER JOIN "
                       + "(SELECT COUNT(*) CountPO, FORM_NAME, PROJECT_ID FROM  PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NOT Null GROUP BY FORM_NAME, PROJECT_ID) Quo "
                       + "ON Quo.PROJECT_ID = tmp.PROJECT_ID AND Quo.FORM_NAME = tmp.FORM_NAME ";
                }
                else
                {
                    // 取得詢價單樣本資訊 - 工資預算 - 圖算數量 * 工資單(預設2500) * 工率*  預算折扣比率
                    sql = "SELECT tmp.*,CountPO,(SELECT SUM(v.QTY * w.RATIO * it.BUDGET_WAGE_RATIO / 100 * @wageunitprice) as BudgetAmount FROM PLAN_ITEM it LEFT JOIN "
                        + "vw_MAP_MATERLIALIST v ON it.PLAN_ITEM_ID = v.PROJECT_ITEM_ID LEFT JOIN TND_WAGE w ON it.PLAN_ITEM_ID = w.PROJECT_ITEM_ID "
                        + "WHERE it.PLAN_ITEM_ID in (SELECT  iit.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY_ITEM iit WHERE iit.INQUIRY_FORM_ID = tmp.INQUIRY_FORM_ID)) AS BudgetAmount "
                        + "FROM(SELECT * FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID is Null AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效')=@status AND ISNULL(ISWAGE, 'N') = 'Y') tmp LEFT OUTER JOIN "
                        + "(SELECT COUNT(*) CountPO, FORM_NAME, PROJECT_ID FROM  PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NOT Null "
                        + "GROUP BY FORM_NAME, PROJECT_ID) Quo ON Quo.PROJECT_ID = tmp.PROJECT_ID AND Quo.FORM_NAME = tmp.FORM_NAME ;";
                    if (null != project.WAGE_MULTIPLIER)
                    {
                        wageunitprice = (decimal)project.WAGE_MULTIPLIER;
                    }
                    parameters.Add(new SqlParameter("wageunitprice", wageunitprice));
                }

                logger.Debug("sql=" + sql + ",projectId=" + project.PROJECT_ID);
                lst = context.Database.SqlQuery<PURCHASE_ORDER>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //取得預算總價
        public void getBudgetSummary(TND_PROJECT project)
        {
            string sql = "SELECT SUM(mBudget) Material_Budget,SUM(wBudget) Wage_Budget FROM ("
                + "SELECT pi.PLAN_ITEM_ID,pi.PROJECT_ID,pi.ITEM_ID,pi.ITEM_DESC,pi.ITEM_QUANTITY,map.QTY mapQty, pi.ITEM_UNIT,"
                + "pi.ITEM_UNIT_PRICE SellProice, pji.ITEM_UNIT_PRICE Cost, isNull(pi.BUDGET_RATIO, 100) BUDGET_RATIO,"
                + "(map.QTY * pji.ITEM_UNIT_PRICE * isNull(pi.BUDGET_RATIO, 100) / 100) mBudget,"
                + "isNull(pi.BUDGET_WAGE_RATIO, 100) BUDGET_WAGE_RATIO,isnull(w.RATIO, 0) wRatio,"
                + "(@wageunitprice * ISNULL(map.QTY, 0) * isNull(pi.BUDGET_WAGE_RATIO, 100) * isnull(w.RATIO, 0) / 100) wBudget "
                + "FROM PLAN_ITEM pi LEFT OUTER JOIN TND_PROJECT_ITEM pji on pi.PLAN_ITEM_ID = pji.PROJECT_ITEM_ID "
                + "LEFT OUTER JOIN  vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID "
                + "LEFT OUTER JOIN TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID "
                + "WHERE pi.PROJECT_ID = @projectid) A; ";
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //設定專案預算資料
                parameters.Add(new SqlParameter("projectid", project.PROJECT_ID));
                //專案工資若未設定則以2500 計算
                if (null == project.WAGE_MULTIPLIER)
                {
                    project.WAGE_MULTIPLIER = 2500;
                }
                parameters.Add(new SqlParameter("wageunitprice", project.WAGE_MULTIPLIER));
                logger.Debug("sql=" + sql + ",projectId=" + project.PROJECT_ID);
                POFormData.BudgetSummary = context.Database.SqlQuery<BUDGET_SUMMANY>(sql, parameters.ToArray()).First();
            }
        }

        public List<PlanSupplierFormFunction> getFormByProject(string projectid, string _status, string _type, string formname)
        {
            string status = "有效";
            if (null != _status && _status != "*")
            {
                status = _status;
            }
            string type = "N";
            if (null != _type && _type != "*")
            {
                type = _type;
            }
            List<PlanSupplierFormFunction> lst = new List<PlanSupplierFormFunction>();
            string sql = "SELECT a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, SUM(b.ITEM_QTY*b.ITEM_UNIT_PRICE) AS TOTAL_PRICE, ROW_NUMBER() OVER(ORDER BY a.INQUIRY_FORM_ID DESC) AS NO, ISNULL(a.STATUS, '有效') AS STATUS, ISNULL(a.ISWAGE,'N') ISWAGE " +
                    "FROM PLAN_SUP_INQUIRY a left JOIN PLAN_SUP_INQUIRY_ITEM b ON a.INQUIRY_FORM_ID = b.INQUIRY_FORM_ID WHERE ISNULL(a.STATUS,'有效')=@status AND ISNULL(a.ISWAGE,'N')=@type  ";
            var parameters = new List<SqlParameter>();
            //設定專案編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            //設定詢價單是否有效
            parameters.Add(new SqlParameter("status", status));
            //設定詢價單為工資或材料
            parameters.Add(new SqlParameter("type", type));
            //詢價單名稱條件
            if (null != formname && formname != "")
            {
                sql = sql + "AND a.FORM_NAME LIKE @formname ";
                parameters.Add(new SqlParameter("formname", "%" + formname + "%"));
            }
            sql = sql + " GROUP BY a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, a.PROJECT_ID, a.STATUS, a.ISWAGE HAVING  a.SUPPLIER_ID IS NOT NULL " +
                    "AND a.PROJECT_ID =@projectid ORDER BY a.INQUIRY_FORM_ID DESC, a.FORM_NAME ;";

            logger.Info("sql=" + sql);
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<PlanSupplierFormFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("get plan supplier form function count:" + lst.Count);
            }
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
                        existItem.ITEM_REMARK = item.ITEM_REMARK;
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

        //更新採購詢價單單價
        public int refreshSupplierFormItem(string formid, List<PLAN_SUP_INQUIRY_ITEM> lstItem)
        {
            logger.Info("Update plan supplier inquiry form id =" + formid);
            int j = 0;
            using (var context = new topmepEntities())
            {
                //將item單價寫入 
                foreach (PLAN_SUP_INQUIRY_ITEM item in lstItem)
                {
                    PLAN_SUP_INQUIRY_ITEM existItem = null;
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("formid", formid));
                    parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                    string sql = "SELECT * FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                    logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                    PLAN_SUP_INQUIRY_ITEM excelItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                    existItem = context.PLAN_SUP_INQUIRY_ITEM.Find(excelItem.INQUIRY_ITEM_ID);
                    logger.Debug("find exist item=" + existItem.ITEM_DESC);
                    existItem.ITEM_UNIT_PRICE = item.ITEM_UNIT_PRICE;
                    existItem.ITEM_REMARK = item.ITEM_REMARK;
                    context.PLAN_SUP_INQUIRY_ITEM.AddOrUpdate(existItem);
                }
                j = context.SaveChanges();
                logger.Debug("Update plan supplier inquiry form item =" + j);
            }
            return j;
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

        //判斷詢價單是否已被寫入PLAN_ITEM(詢價單已為發包採用)
        public Boolean getSupplierContractByFormId(string formid)
        {
            logger.Info("get boolean of formid in the plan item by formid=" + formid);
            //處理SQL 預先填入ID,設定集合處理參數
            Boolean count = false;
            using (var context = new topmepEntities())
            {
                count = context.Database.SqlQuery<Boolean>("SELECT CAST(COUNT(*) AS BIT) AS BOOLEAN FROM PLAN_ITEM WHERE INQUIRY_FORM_ID =@formid OR MAN_FORM_ID =@formid  ; "
            , new SqlParameter("formid", formid)).FirstOrDefault();
            }

            return count;
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
                lst = context.Database.SqlQuery<plansummary>("SELECT  A.PROJECT_ID + A.ID + A.FORM_NAME AS CONTRACT_ID, A.SUPPLIER_ID, A.FORM_NAME, " +
                    "SUM(A.MapQty * A.ITEM_UNIT_COST) MATERIAL_COST, SUM(A.MapQty * ISNULL(A.MAN_PRICE, 0)) WAGE_COST, " +
                    "SUM(A.ITEM_QUANTITY * ISNULL(A.ITEM_UNIT_PRICE, 0)) REVENUE, SUM(A.MapQty * A.TndFormPrice * ISNULL(A.BUDGET_RATIO, 100) / 100) BUDGET, " +
                    "(SUM(A.MapQty * A.ITEM_UNIT_COST) + SUM(A.MapQty * ISNULL(A.MAN_PRICE, 0))) COST, (SUM(A.ITEM_QUANTITY * ISNULL(A.ITEM_UNIT_PRICE, 0)) - " +
                    "SUM(A.MapQty * A.ITEM_UNIT_COST) - SUM(A.MapQty * ISNULL(A.MAN_PRICE, 0))) PROFIT, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY A.SUPPLIER_ID) AS NO FROM(SELECT pi.*, s.SUPPLIER_ID AS ID, map.QTY AS MapQty, " +
                    "tpi.ITEM_UNIT_PRICE AS TndFormPrice FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                    "pi.SUPPLIER_ID = s.COMPANY_NAME LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT_ITEM tpi ON pi.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID  " +
                    ")A WHERE A.PROJECT_ID = @projectid AND A.ITEM_UNIT_COST IS NOT NULL " +
                    "GROUP BY A.PROJECT_ID, A.ID, A.FORM_NAME, A.SUPPLIER_ID ; "
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
                lst = context.Database.SqlQuery<plansummary>("SELECT  A.PROJECT_ID + A.ID + A.MAN_FORM_NAME AS CONTRACT_ID, A.MAN_SUPPLIER_ID, A.MAN_FORM_NAME, " +
                    "SUM(A.MapQty * ISNULL(A.MAN_PRICE, 0)) WAGE_COST, SUM(A.MapQty * A.Ratio * A.WAGE_MULTIPLIER * ISNULL(A.BUDGET_WAGE_RATIO, 100) / 100) AS WAGE_BUDGET, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY A.MAN_SUPPLIER_ID) AS NO FROM(SELECT pi.*, map.QTY AS MapQty, w.RATIO AS Ratio, p.WAGE_MULTIPLIER, " +
                    "s.SUPPLIER_ID AS ID FROM PLAN_ITEM pi LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT OUTER JOIN TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID " +
                    "LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID LEFT JOIN TND_SUPPLIER s ON " +
                    "pi.MAN_SUPPLIER_ID = s.COMPANY_NAME)A WHERE A.PROJECT_ID = 'P00023' and A.MAN_PRICE IS NOT NULL " +
                    "GROUP BY A.PROJECT_ID, A.MAN_SUPPLIER_ID, A.MAN_FORM_NAME, A.ID  ; "
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
                    "p.SUPPLIER_ID, p.FORM_NAME, sum(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) REVENUE, " +
                    "sum(map.QTY * tpi.ITEM_UNIT_PRICE * ISNULL(p.BUDGET_RATIO,100) / 100) + sum(map.QTY * w.RATIO * tp.WAGE_MULTIPLIER * ISNULL(p.BUDGET_WAGE_RATIO,100) / 100) BUDGET, " +
                    "(sum(map.QTY * p.ITEM_UNIT_COST) + SUM(map.QTY * ISNULL(p.MAN_PRICE, 0))) COST, (sum(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) - sum(map.QTY * p.ITEM_UNIT_COST) - SUM(map.QTY * ISNULL(p.MAN_PRICE, 0))) PROFIT, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY p.SUPPLIER_ID) AS NO FROM PLAN_ITEM p LEFT JOIN vw_MAP_MATERLIALIST map ON p.PLAN_ITEM_ID = map.PROJECT_ITEM_ID " +
                    "LEFT JOIN TND_PROJECT_ITEM tpi ON p.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID LEFT JOIN TND_WAGE w ON p.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT tp ON p.PROJECT_ID = tp.PROJECT_ID WHERE p.PROJECT_ID = @projectid " +
                    "GROUP BY p.PROJECT_ID, p.SUPPLIER_ID, p.FORM_NAME)A ; "
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
                         "AND p.ITEM_UNIT_COST IS NOT NULL GROUP BY p.FORM_NAME))C LEFT OUTER JOIN " +
                         "(select  B.type, B.INQUIRY_FORM_ID, B.TOTAL_ROW AS TOTAL_ROWS, B.TOTALPRICE AS TOTAL_PRICE FROM (select p.FORM_NAME as type, p.INQUIRY_FORM_ID " +
                         "from PLAN_SUP_INQUIRY_ITEM pi LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL " +
                         "and pi.ITEM_UNIT_PRICE is not null GROUP BY p.INQUIRY_FORM_ID, p.FORM_NAME HAVING p.FORM_NAME NOT IN " +
                         "(SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.ITEM_UNIT_COST IS NOT NULL GROUP BY p.FORM_NAME)) A " +
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
                         "AND p.MAN_PRICE IS NOT NULL GROUP BY p.FORM_NAME))C LEFT OUTER JOIN " +
                         "(select  B.type, B.INQUIRY_FORM_ID, B.TOTAL_ROW AS TOTAL_ROWS, B.TOTALPRICE AS TOTAL_PRICE FROM (select p.FORM_NAME as type, p.INQUIRY_FORM_ID " +
                         "from PLAN_SUP_INQUIRY_ITEM pi LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL " +
                         "and pi.ITEM_UNIT_PRICE is not null GROUP BY p.INQUIRY_FORM_ID, p.FORM_NAME HAVING p.FORM_NAME NOT IN " +
                         "(SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.MAN_PRICE IS NOT NULL GROUP BY p.FORM_NAME)) A " +
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
                "f.SUPPLIER_ID as SUPPLIER_NAME, f.FORM_NAME AS FORM_NAME, ISNULL(f.STATUS,'有效') STATUS, SUM(pfitem.ITEM_UNIT_PRICE*pfitem.ITEM_QTY) as TAmount, " +
                "ISNULL(CEILING(SUM(pfitem.ITEM_UNIT_PRICE*pfitem.ITEM_QTY) / SUM(w.RATIO*map.QTY*ISNULL(p.WAGE_MULTIPLIER, 0))),0) as AvgMPrice " +
                "FROM PLAN_ITEM pItem LEFT OUTER JOIN " +
                "PLAN_SUP_INQUIRY_ITEM pfItem ON pItem.PLAN_ITEM_ID = pfItem.PLAN_ITEM_ID " +
                "inner join PLAN_SUP_INQUIRY f on pfItem.INQUIRY_FORM_ID = f.INQUIRY_FORM_ID " +
                "left join TND_WAGE w on pItem.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT JOIN vw_MAP_MATERLIALIST map ON pItem.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT p ON pItem.PROJECT_ID = p.PROJECT_ID " +
                "WHERE pItem.PROJECT_ID = @projectid AND f.SUPPLIER_ID is not null AND ISNULL(f.STATUS,'有效') <> '註銷' AND ISNULL(f.ISWAGE,'N')=@iswage  ";
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
                "pitem.ITEM_UNIT_COST 材料單價, " +
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
                "pitem.ITEM_UNIT_COST 材料單價, fitem.ITEM_UNIT_PRICE " +
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
            string sql = "UPDATE PLAN_ITEM SET ITEM_UNIT_COST =@price WHERE PLAN_ITEM_ID=@pitemid ";
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
        //將發包廠商之詢價單價格寫入PLAN_ITEM且寫入後不得再覆寫
        public int batchUpdateCostFromQuote(string formid, string iswage)
        {
            int i = 0;
            logger.Info("Copy cost from Quote to Tnd by form id" + formid);
            string sql = "UPDATE  PLAN_ITEM SET item_unit_cost = i.ITEM_UNIT_PRICE, supplier_id = i.SUPPLIER_ID, form_name = i.FORM_NAME, inquiry_form_id = i.INQUIRY_FORM_ID " +
                "FROM(select i.plan_item_id, fi.ITEM_UNIT_PRICE, fi.INQUIRY_FORM_ID, pf.SUPPLIER_ID, pf.FORM_NAME from PLAN_ITEM i " +
                ", PLAN_SUP_INQUIRY_ITEM fi, PLAN_SUP_INQUIRY pf " +
               "where i.PLAN_ITEM_ID = fi.PLAN_ITEM_ID and fi.INQUIRY_FORM_ID = pf.INQUIRY_FORM_ID and fi.INQUIRY_FORM_ID = @formid) i " +
                "WHERE  i.plan_item_id = PLAN_ITEM.PLAN_ITEM_ID AND PLAN_ITEM.ITEM_UNIT_COST IS NULL ";

            //將工資報價單更新工資報價欄位
            if (iswage == "Y")
            {
                sql = "UPDATE  PLAN_ITEM SET man_price = i.ITEM_UNIT_PRICE, man_supplier_id = i.SUPPLIER_ID, man_form_name = i.FORM_NAME, man_form_id = i.INQUIRY_FORM_ID "
                + "FROM (select i.plan_item_id, fi.ITEM_UNIT_PRICE, fi.INQUIRY_FORM_ID, pf.SUPPLIER_ID, pf.FORM_NAME from PLAN_ITEM i "
                + ", PLAN_SUP_INQUIRY_ITEM fi, PLAN_SUP_INQUIRY pf "
                + "where i.PLAN_ITEM_ID = fi.PLAN_ITEM_ID and fi.INQUIRY_FORM_ID = pf.INQUIRY_FORM_ID and fi.INQUIRY_FORM_ID = @formid) i "
                + "WHERE  i.plan_item_id = PLAN_ITEM.PLAN_ITEM_ID AND PLAN_ITEM.MAN_PRICE IS NULL ";
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
            //將材料合約編號寫入PLAN PAYMENT TERMS
            logger.Info("copy contract id into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                List<topmeperp.Models.PLAN_PAYMENT_TERMS> lstItem = new List<PLAN_PAYMENT_TERMS>();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID, TYPE) " +
                       "SELECT distinct ('" + projectid + "' + A.SUPPLIER_ID + A.FORM_NAME) AS contractid, '" + projectid + "', 'S' FROM " +
                       "(SELECT pi.PROJECT_ID, pi.SUPPLIER_ID AS SUPPLIER_NAME, pi.FORM_NAME, s.SUPPLIER_ID FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                       "pi.SUPPLIER_ID = s.COMPANY_NAME WHERE pi.SUPPLIER_ID IS NOT NULL)A WHERE A.PROJECT_ID = '" + projectid + "' " +
                       "AND '" + projectid + "' + A.SUPPLIER_ID + A.FORM_NAME NOT IN(SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";

                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }
        //將工資合約編號寫入PLAN PAYMENT TERMS
        public int addContractIdForWage(string projectid)
        {
            int i = 0;
            logger.Info("copy contract id from wage into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                List<topmeperp.Models.PLAN_PAYMENT_TERMS> lstItem = new List<PLAN_PAYMENT_TERMS>();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID, TYPE) " +
                   "SELECT distinct ('" + projectid + "' + A.SUPPLIER_ID + A.MAN_FORM_NAME) AS contractid, '" + projectid + "', 'S' FROM " +
                   "(SELECT pi.PROJECT_ID, pi.MAN_SUPPLIER_ID AS SUPPLIER_NAME, pi.MAN_FORM_NAME, s.SUPPLIER_ID FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                   "pi.MAN_SUPPLIER_ID = s.COMPANY_NAME WHERE pi.MAN_SUPPLIER_ID IS NOT NULL)A WHERE A.PROJECT_ID = '" + projectid + "' " +
                   "AND '" + projectid + "' + A.SUPPLIER_ID + A.MAN_FORM_NAME NOT IN (SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";

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
                lst = context.Database.SqlQuery<PLAN_ITEM>("SELECT * FROM PLAN_ITEM WHERE PROJECT_ID =@projectid AND ITEM_UNIT IS NOT NULL AND ITEM_UNIT <> '' AND ITEM_UNIT_COST IS NULL ;"
                    , new SqlParameter("projectid", projectid)).ToList();
            }
            logger.Info("get plan pending items count:" + lst.Count);
            return lst;
        }


        public List<PLAN_ITEM> getPendingItems4Wage(string projectid)
        {
            List<PLAN_ITEM> lst = new List<PLAN_ITEM>();
            using (var context = new topmepEntities())
            {
                // ITEM_UNIT IS NOT NULL(確認單位欄位是空值就是不需採購的欄位嗎) AND ITEM_UNIT_PRICE IS NULL
                lst = context.Database.SqlQuery<PLAN_ITEM>("SELECT * FROM PLAN_ITEM WHERE PROJECT_ID =@projectid AND ITEM_UNIT IS NOT NULL AND ITEM_UNIT <> '' AND MAN_PRICE IS NULL ;"
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

        #region 物料管理之進銷存
        public PLAN_PURCHASE_REQUISITION formPR = null;
        public List<PurchaseRequisition> PRItem = null;
        public List<PurchaseRequisition> DOItem = null;

        public List<PurchaseRequisition> getPurchaseItemByMap(string projectid, List<string> lstItemId)
        {
            //取得任務採購內容
            logger.Info("get plan item by map ");
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            using (var context = new topmepEntities())
            {
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

                string sql = "SELECT pi.* , map.QTY AS MAP_QTY, B.CUMULATIVE_QTY, C.ALL_RECEIPT_QTY- D.DELIVERY_QTY AS INVENTORY_QTY FROM PLAN_ITEM pi  " +
                    "JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN (SELECT pri.PLAN_ITEM_ID, SUM(pri.ORDER_QTY) AS CUMULATIVE_QTY " +
                    "FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE PR_ID LIKE 'PPO%' GROUP BY pri.PLAN_ITEM_ID )B ON pi.PLAN_ITEM_ID = B.PLAN_ITEM_ID " +
                    "LEFT JOIN(SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS ALL_RECEIPT_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE PR_ID LIKE 'RP%' GROUP BY " +
                    "pri.PLAN_ITEM_ID)C ON pi.PLAN_ITEM_ID = C.PLAN_ITEM_ID LEFT JOIN (SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid " +
                    "GROUP BY pid.PLAN_ITEM_ID)D ON pi.PLAN_ITEM_ID = D.PLAN_ITEM_ID WHERE pi.PROJECT_ID = @projectid AND pi.PLAN_ITEM_ID IN (" + ItemId + ") ";

                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
                logger.Info("Get task material Info Record Count=" + lstItem.Count);
            }
            return lstItem;
        }

        // 寫入任務採購內容
        public string newPR(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId)
        {
            //1.建立申購單
            logger.Info("create new purchase requisition ");
            string sno_key = "PR";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new purchase requisition =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase Requisition=" + i);
                logger.Info("plan purchase requisition id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
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

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID  "
                + "FROM (SELECT pi.PLAN_ITEM_ID FROM PLAN_ITEM pi WHERE pi.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }
        //更新申購數量
        public int refreshPR(string formid, PLAN_PURCHASE_REQUISITION form, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update plan purchase requisition id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan purchase requisition =" + i);
                    logger.Info("purchase requisition item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.NEED_QTY = item.NEED_QTY;
                        existItem.NEED_DATE = item.NEED_DATE;
                        existItem.REMARK = item.REMARK;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase requisition item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase requisition id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //取得申購單資料
        public List<PRFunction> getPRByPrjId(string projectid, string date, string taskname, string prid, int status)
        {
            logger.Info("search purchase requisition by 申購日期 =" + date + ", 申購單編號 =" + prid + ", 任務名稱 =" + taskname + ", 申購單狀態 =" + status);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            if (10 == status)
            {
                string sql = "SELECT CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.PR_ID, A.STATUS, A.TASK_NAME, B.PR_ID AS CHILD_PR_ID, ROW_NUMBER() OVER(ORDER BY A.PR_ID) AS NO " +
                    "FROM (SELECT pr.CREATE_DATE, pr.PR_ID, pr.PRJ_UID, pt.TASK_NAME, pr.STATUS FROM PLAN_PURCHASE_REQUISITION pr LEFT OUTER JOIN PLAN_TASK pt " +
                    "ON pr.PRJ_UID = pt.PRJ_UID WHERE pr.PROJECT_ID=@projectid AND pr.SUPPLIER_ID IS NULL)A LEFT JOIN (SELECT * FROM PLAN_PURCHASE_REQUISITION pr WHERE pr.PR_ID LIKE 'PPO%')B ON A.PR_ID = B.PARENT_PR_ID ";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                sql = sql + "WHERE A.STATUS = 10 ";

                //申購日期查詢條件
                if (null != date && date != "")
                {
                    DateTime dt = Convert.ToDateTime(date);
                    string DateString = dt.AddDays(1).ToString("yyyy/MM/dd");
                    sql = sql + "AND CREATE_DATE >=@date AND  CREATE_DATE < '" + DateString + "' ";
                    parameters.Add(new SqlParameter("date", date));
                }
                //申購單編號條件
                if (null != prid && prid != "")
                {
                    sql = sql + "AND A.PR_ID =@prid ";
                    parameters.Add(new SqlParameter("prid", prid));
                }
                //任務名稱條件
                if (null != taskname && taskname != "")
                {
                    sql = sql + "AND A.TASK_NAME LIKE @taskname ";
                    parameters.Add(new SqlParameter("taskname", '%' + taskname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get purchase requisition sql=" + sql);
                    lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get purchase requisition count=" + lstForm.Count);
            }
            else
            {
                string sql = "SELECT CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.PR_ID, A.STATUS, A.TASK_NAME, ROW_NUMBER() OVER(ORDER BY A.PR_ID) AS NO " +
                    "FROM (SELECT pr.CREATE_DATE, pr.PR_ID, pr.PRJ_UID, pt.TASK_NAME, pr.STATUS FROM PLAN_PURCHASE_REQUISITION pr LEFT JOIN PLAN_TASK pt " +
                    "ON pr.PRJ_UID = pt.PRJ_UID WHERE pr.PROJECT_ID=@projectid AND pr.SUPPLIER_ID IS NULL)A ";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                sql = sql + "WHERE A.STATUS = 0 ";

                using (var context = new topmepEntities())
                {
                    logger.Debug("get purchase requisition sql=" + sql);
                    lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get purchase requisition count=" + lstForm.Count);
            }
            return lstForm;
        }

        //取得申購單
        public void getPRByPrId(string prid, string parentId)
        {
            logger.Info("get form : formid=" + prid);
            using (var context = new topmepEntities())
            {
                //取得申購單檔頭資訊
                string sql = "SELECT PR_ID, PROJECT_ID, RECIPIENT, LOCATION, PRJ_UID, CREATE_USER_ID, CREATE_DATE, REMARK, SUPPLIER_ID, MODIFY_DATE, PARENT_PR_ID, STATUS, MEMO, MESSAGE, CAUTION FROM " +
                    "PLAN_PURCHASE_REQUISITION WHERE PR_ID =@prid ";
                formPR = context.PLAN_PURCHASE_REQUISITION.SqlQuery(sql, new SqlParameter("prid", prid)).First();
                //取得申購單明細
                PRItem = context.Database.SqlQuery<PurchaseRequisition>("SELECT pri.NEED_QTY, CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, pri.REMARK, pri.PR_ITEM_ID, pri.ORDER_QTY, pri.PLAN_ITEM_ID, pri.RECEIPT_QTY, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.ITEM_FORM_QUANTITY, pi.SYSTEM_MAIN, md.QTY AS MAP_QTY,  " +
                    "B.CUMULATIVE_QTY, C.ALL_RECEIPT_QTY, C.ALL_RECEIPT_QTY - D.DELIVERY_QTY AS INVENTORY_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN TND_MAP_DEVICE md " +
                    "ON pi.PLAN_ITEM_ID = md.PROJECT_ITEM_ID LEFT JOIN (SELECT pri.PLAN_ITEM_ID, SUM(pri.ORDER_QTY) AS CUMULATIVE_QTY " +
                    "FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE PR_ID LIKE 'PPO%' GROUP BY pri.PLAN_ITEM_ID)B ON pri.PLAN_ITEM_ID = B.PLAN_ITEM_ID " +
                    "LEFT JOIN(SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS ALL_RECEIPT_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION ppr " +
                    "ON pri.PR_ID = ppr.PR_ID WHERE ppr.PARENT_PR_ID =@parentId GROUP BY " +
                    "pri.PLAN_ITEM_ID)C ON pri.PLAN_ITEM_ID = C.PLAN_ITEM_ID LEFT JOIN (SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid " +
                    "GROUP BY pid.PLAN_ITEM_ID)D ON pri.PLAN_ITEM_ID = D.PLAN_ITEM_ID WHERE PR_ID =@prid", new SqlParameter("prid", prid), new SqlParameter("parentId", parentId)).ToList();

                logger.Debug("get purchase requisition item count:" + PRItem.Count);
                //取得領料明細
                DOItem = context.Database.SqlQuery<PurchaseRequisition>("SELECT pid.PLAN_ITEM_ID, pid.DELIVERY_QTY, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN " +
                    "FROM PLAN_ITEM_DELIVERY pid LEFT JOIN PLAN_ITEM pi ON pid.PLAN_ITEM_ID = pi.PLAN_ITEM_ID WHERE pid.DELIVERY_ORDER_ID =@prid", new SqlParameter("prid", prid)).ToList();

                logger.Debug("get delivery item count:" + DOItem.Count);
            }
        }

        public string getParentPrIdByPrId(string prid)
        {
            string parentid = null;
            using (var context = new topmepEntities())
            {
                parentid = context.Database.SqlQuery<string>("select DISTINCT PARENT_PR_ID FROM PLAN_PURCHASE_REQUISITION WHERE PARENT_PR_ID =@prid  "
               , new SqlParameter("prid", prid)).FirstOrDefault();
            }
            return parentid;
        }
        public PLAN_PURCHASE_REQUISITION table = null;
        //更新申購單資料
        public int updatePR(string formid, PLAN_PURCHASE_REQUISITION pr, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update purchase requisition id =" + formid);
            table = pr;
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(table).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update purchase requisition =" + i);
                    logger.Info("purchase requisition item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        logger.Debug("purchase requisition item id=" + item.PR_ITEM_ID);
                        if (item.PR_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(item.PR_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                            PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.NEED_QTY = item.NEED_QTY;
                        existItem.NEED_DATE = item.NEED_DATE;
                        existItem.REMARK = item.REMARK;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase requisition item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase requisition id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        //取得申購單by供應商
        public List<PurchaseOrderFunction> getPRBySupplier(string projectid)
        {
            List<PurchaseOrderFunction> lstPO = new List<PurchaseOrderFunction>();
            using (var context = new topmepEntities())
            {
                string sql = "SELECT B.KEYNAME, CONVERT(char(10), B.CREATE_DATE, 111) AS CREATE_DATE, B.PR_ID, B.SUPPLIER_ID, MIN(CONVERT(char(10), B.NEED_DATE, 111)) AS NEED_DATE, " +
                    "B.PROJECT_ID FROM (SELECT DISTINCT(A.PROJECT_ID + '-' + PR_ID + '-' + SUPPLIER_ID + '-' + CONVERT(char(10), NEED_DATE, 111)) AS NAME, " +
                    "A.PROJECT_ID + '-' + PR_ID + '-' + ISNULL(SUPPLIER_ID,'') AS KEYNAME, CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.PR_ID, A.SUPPLIER_ID, A.PROJECT_ID, " +
                    "A.NEED_DATE FROM (SELECT pri.*, pi.ITEM_ID, pi.ITEM_DESC, pi.SUPPLIER_ID, pr.CREATE_DATE, pr.PROJECT_ID FROM PLAN_PURCHASE_REQUISITION_ITEM pri " +
                    "JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID =@projectid " +
                    "AND pr.SUPPLIER_ID IS NULL AND pr.STATUS > 0 AND pr.PR_ID LIKE 'PR%' )A WHERE A.PR_ID + A.SUPPLIER_ID NOT IN (SELECT DISTINCT(pr.PARENT_PR_ID + pr.SUPPLIER_ID) AS ORDER_RECORD " +
                    "FROM PLAN_PURCHASE_REQUISITION pr WHERE pr.PARENT_PR_ID + pr.SUPPLIER_ID IS NOT NULL))B GROUP BY B.KEYNAME, CONVERT(char(10), B.CREATE_DATE, 111), " +
                    "B.PR_ID, B.SUPPLIER_ID, B.PROJECT_ID ORDER BY NEED_DATE ";

                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstPO = context.Database.SqlQuery<PurchaseOrderFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("Get Purchase Requisition By Suplier Record Count =" + lstPO.Count);
            }
            return lstPO;
        }
        //取得申購單項目by供應商
        public List<PurchaseRequisition> getPurchaseItemBySupplier(string id)
        {
            //取得各供應商採購內容
            logger.Info("get purchase requisition item by supplier ");
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            using (var context = new topmepEntities())
            {
                string sql = "SELECT pi.PLAN_ITEM_ID, pri.PR_ITEM_ID, pri.NEED_QTY, CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, pri.REMARK , pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.ITEM_FORM_QUANTITY, " +
                    "pi.SUPPLIER_ID, B.CUMULATIVE_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi on pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                    "LEFT JOIN (SELECT pri.PLAN_ITEM_ID, SUM(pri.ORDER_QTY) AS CUMULATIVE_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE PR_ID LIKE 'PR%' GROUP BY pri.PLAN_ITEM_ID)B " +
                    "ON pri.PLAN_ITEM_ID = B.PLAN_ITEM_ID WHERE pri.PR_ID + '-' + pi.SUPPLIER_ID =@id ";

                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", id));
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
                logger.Info("Get purchase requisition item by supplier Record Count=" + lstItem.Count);
            }
            return lstItem;
        }

        // 寫入採購內容
        public string newPO(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId, string parentid)
        {
            //1.建立採購單
            logger.Info("create new purchase order ");
            string sno_key = "PPO";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new purchase order =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase Order=" + i);
                logger.Info("plan purchase Order id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
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

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID, NEED_QTY, NEED_DATE, REMARK) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.NEED_QTY as NEED_QTY, A.NEED_DATE as NEED_DATE, A.REMARK as REMARK  "
                + "FROM (SELECT pri.* FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE pri.PR_ID = '" + parentid + "' AND pri.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }
        //更新採購數量
        public int refreshPO(string formid, PLAN_PURCHASE_REQUISITION form, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update plan purchase order id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan purchase order =" + i);
                    logger.Info("purchase order item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.ORDER_QTY = item.ORDER_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase order item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase order id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //取得採購單資料
        public List<PRFunction> getPOByPrjId(string projectid, string date, string supplier, string prid, string parentPrid)
        {

            logger.Info("search purchase order by 採購日期 =" + date + ", 採購單編號 =" + prid + ", 供應商名稱 =" + supplier + ", 申購單編號 =" + parentPrid);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, SUPPLIER_ID, PARENT_PR_ID, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE PROJECT_ID =@projectid AND SUPPLIER_ID IS NOT NULL AND PR_ID NOT LIKE 'RP%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //採購日期查詢條件
            if (null != date && date != "")
            {
                DateTime dt = Convert.ToDateTime(date);
                string DateString = dt.AddDays(1).ToString("yyyy/MM/dd");
                sql = sql + "AND CREATE_DATE >=@date AND  CREATE_DATE < '" + DateString + "' ";
                parameters.Add(new SqlParameter("date", date));
            }
            //採購單編號條件
            if (null != prid && prid != "")
            {
                sql = sql + "AND PR_ID =@prid ";
                parameters.Add(new SqlParameter("prid", prid));
            }
            //供應商條件
            if (null != supplier && supplier != "")
            {
                sql = sql + "AND SUPPLIER_ID LIKE @supplier ";
                parameters.Add(new SqlParameter("supplier", '%' + supplier + '%'));
            }
            //申購單編號條件
            if (null != parentPrid && parentPrid != "")
            {
                sql = sql + "AND PARENT_PR_ID =@parentPrid ";
                parameters.Add(new SqlParameter("parentPrid", parentPrid));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase order sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase order count=" + lstForm.Count);
            return lstForm;
        }

        //更新採購單資料
        public int updatePO(string formid, PLAN_PURCHASE_REQUISITION pr, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update purchase order id =" + formid);
            table = pr;
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(table).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update purchase order =" + i);
                    logger.Info("purchase order item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        logger.Debug("purchase order item id=" + item.PR_ITEM_ID);
                        if (item.PR_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(item.PR_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                            PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.ORDER_QTY = item.ORDER_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase order item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase order id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        //更新採購單memo
        public int changeMemo(string formid, string memo)
        {
            int i = 0;
            logger.Info("Update PO memo, it's formid=" + formid + ", memo =" + memo);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_PURCHASE_REQUISITION SET MEMO=@memo, MODIFY_DATE =@datetime  WHERE PR_ID=@formid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("memo", memo));
            parameters.Add(new SqlParameter("formid", formid));
            parameters.Add(new SqlParameter("datetime", DateTime.Now));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Debug("Update PO memo :" + i);
            return i;
        }
        // 寫入驗收內容
        public string newRP(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId, string parentid)
        {
            //1.建立驗收資料
            logger.Info("create new receipt ");
            string sno_key = "RP";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new purchase receipt =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase receipt=" + i);
                logger.Info("plan purchase receipt id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
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

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID, NEED_QTY, NEED_DATE, REMARK, ORDER_QTY) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.NEED_QTY as NEED_QTY, A.NEED_DATE as NEED_DATE, A.REMARK as REMARK, A.ORDER_QTY as ORDER_QTY  "
                + "FROM (SELECT pri.* FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE pri.PR_ID = '" + parentid + "' AND pri.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }

        //更新驗收數量
        public int refreshRP(string formid, PLAN_PURCHASE_REQUISITION form, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update plan purchase receipt id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan purchase receipt =" + i);
                    logger.Info("purchase receipt item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.RECEIPT_QTY = item.RECEIPT_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase reeipt item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase receipt id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //更新驗收單資料
        public int updateRP(string formid, PLAN_PURCHASE_REQUISITION pr, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update purchase receipt id =" + formid);
            table = pr;
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(table).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update purchase receipt =" + i);
                    logger.Info("purchase receipt item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        logger.Debug("purchase receipt item id=" + item.PR_ITEM_ID);
                        if (item.PR_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(item.PR_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                            PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.RECEIPT_QTY = item.RECEIPT_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase receipt item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase receipt id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //取得驗收單資料
        public List<PRFunction> getRPByPrId(string prid)
        {

            logger.Info("search purchase receipt by 採購單編號 =" + prid);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, SUPPLIER_ID, PR_ID + '-' + PARENT_PR_ID AS ALL_KEY, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE SUPPLIER_ID IS NOT NULL AND PARENT_PR_ID =@prid AND PR_ID LIKE 'RP%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prid", prid));

            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase receipt sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase receipt count=" + lstForm.Count);
            return lstForm;
        }

        //取得驗收單資料以供領料使用
        public List<PRFunction> getRP4Delivery(string prjid, string keyword)
        {

            logger.Info("search receipt for delivery by 專案編號 =" + prjid + ", 關鍵字名稱 =" + keyword);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, SUPPLIER_ID, REMARK, MEMO, MESSAGE, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE PROJECT_ID =@prjid AND PR_ID LIKE 'RP%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prjid", prjid));

            //關鍵字條件
            if (null != keyword && keyword != "")
            {
                sql = sql + "AND MEMO LIKE @keyword OR PROJECT_ID =@prjid AND PR_ID LIKE 'RP%' AND MESSAGE LIKE @keyword OR PROJECT_ID =@prjid AND PR_ID LIKE 'RP%' AND REMARK LIKE @keyword ";
                parameters.Add(new SqlParameter("keyword", '%' + keyword + '%'));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get receipt for delivery sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get receipt for delivery count=" + lstForm.Count);
            return lstForm;
        }
        //取得物料庫存數量
        public List<PurchaseRequisition> getInventoryByPrjId(string prjid, string itemName, string typeMain, string typeSub, string systemMain, string systemSub)
        {

            logger.Info("search inventory by 專案編號 =" + prjid + ", 物料名稱 =" + itemName + ", 九宮格 =" + typeMain + ", 次九宮格 =" + typeSub + ", 主系統 =" + systemMain + ", 次系統 =" + systemSub);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pri.PLAN_ITEM_ID, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN, SUM(pri.RECEIPT_QTY) - ISNULL(A.DELIVERY_QTY, 0) AS INVENTORY_QTY " +
                "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                "LEFT JOIN (SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid " +
                "GROUP BY pid.PLAN_ITEM_ID)A ON pri.PLAN_ITEM_ID = A.PLAN_ITEM_ID GROUP BY pri.PLAN_ITEM_ID, A.DELIVERY_QTY, " +
                "pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.TYPE_CODE_1, pi.TYPE_CODE_2, pi.SYSTEM_MAIN, pi.SYSTEM_SUB, pi.EXCEL_ROW_ID HAVING pri.PLAN_ITEM_ID IN (SELECT pi.PLAN_ITEM_ID FROM PLAN_ITEM pi WHERE pi.PROJECT_ID =@prjid) ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prjid", prjid));

            //物料名稱條件
            if (null != itemName && itemName != "")
            {
                sql = sql + "AND pi.ITEM_DESC LIKE @itemName ";
                parameters.Add(new SqlParameter("itemName", '%' + itemName + '%'));
            }
            //九宮格條件
            if (null != typeMain && typeMain != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 =@typeMain ";
                parameters.Add(new SqlParameter("typeMain", typeMain));
            }
            //次九宮格條件
            if (null != typeSub && typeSub != "")
            {
                sql = sql + "AND pi.TYPE_CODE_2 =@typeSub ";
                parameters.Add(new SqlParameter("typeSub", typeSub));
            }
            //主系統條件
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND REPLACE(pi.SYSTEM_MAIN,' ','') =@systemMain ";
                parameters.Add(new SqlParameter("systemMain", systemMain));
            }
            //次系統條件
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND REPLACE(pi.SYSTEM_SUB,' ','') =@systemSub ";
                parameters.Add(new SqlParameter("systemSub", systemSub));
            }
            sql = sql + "ORDER BY pi.EXCEL_ROW_ID ASC ";
            using (var context = new topmepEntities())
            {
                logger.Debug("get inventory sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get inventory count=" + lstItem.Count);
            return lstItem;
        }

        // 寫入領料內容
        public string newDelivery(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId, string createid)
        {
            //1.新增領料品項
            logger.Info("create new delivery item ");
            string sno_key = "DO";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new delivery form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Delivery Form=" + i);
                logger.Info("plan delivery form id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_ITEM_DELIVERY> lstItem = new List<PLAN_ITEM_DELIVERY>();
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

                string sql = "INSERT INTO PLAN_ITEM_DELIVERY (DELIVERY_ORDER_ID, PLAN_ITEM_ID, PROJECT_ID, CREATE_USER_ID) "
                + "SELECT '" + form.PR_ID + "' as DELIVERY_ORDER_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, '" + projectid + "' as PROJECT_ID, '" + createid + "' as CREATE_USER_ID  "
                + "FROM (SELECT pi.PLAN_ITEM_ID FROM PLAN_ITEM pi WHERE pi.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }
        //新增領料數量
        public int refreshDelivery(string deliveryorderid, List<PLAN_ITEM_DELIVERY> lstItem)
        {
            logger.Info("Update delivery items, it's delivery order id =" + deliveryorderid);
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    //將item資料寫入 
                    foreach (PLAN_ITEM_DELIVERY item in lstItem)
                    {
                        PLAN_ITEM_DELIVERY existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", deliveryorderid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ITEM_DELIVERY WHERE DELIVERY_ORDER_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + deliveryorderid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_ITEM_DELIVERY excelItem = context.PLAN_ITEM_DELIVERY.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ITEM_DELIVERY.Find(excelItem.DELIVERY_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.DELIVERY_QTY = item.DELIVERY_QTY;
                        existItem.CREATE_DATE = DateTime.Now;
                        context.PLAN_ITEM_DELIVERY.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update delivery item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new delivery id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return j;
        }

        //取得物料進出紀錄
        public List<PurchaseRequisition> getDeliveryByItemId(string itemid)
        {

            logger.Info(" get receipt record and delivery record by 物料編號 =" + itemid);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT A.* FROM (SELECT pid.DELIVERY_ORDER_ID, pid.CREATE_DATE, pid.DELIVERY_QTY, pr.PARENT_PR_ID, ROW_NUMBER() OVER(ORDER BY pid.CREATE_DATE ASC) AS NO " +
                "FROM PLAN_ITEM_DELIVERY pid LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pid.DELIVERY_ORDER_ID = pr.PR_ID WHERE pid.PLAN_ITEM_ID =@itemid " +
                "UNION SELECT pri.PR_ID, pr.CREATE_DATE, pri.RECEIPT_QTY, pr.PARENT_PR_ID, ROW_NUMBER() OVER(ORDER BY pr.CREATE_DATE ASC) AS NO FROM PLAN_PURCHASE_REQUISITION_ITEM pri " +
                "LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PR_ID LIKE 'RP%' AND pri.PLAN_ITEM_ID =@itemid)A ORDER BY A.CREATE_DATE ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("itemid", itemid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get receipt and delivery sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get receipt record and delivery record count=" + lstItem.Count);
            return lstItem;
        }

        //寫入領料內容
        public string newDO(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId)
        {
            //1.建立領料單
            logger.Info("create new delivery form ");
            string sno_key = "DF";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new delivery form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add delivery form=" + i);
                logger.Info("plan delivery form id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
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

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID, RECEIPT_QTY, NEED_DATE, REMARK, ORDER_QTY) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.RECEIPT_QTY as RECEIPT_QTY, A.NEED_DATE as NEED_DATE, A.REMARK as REMARK, A.ORDER_QTY as ORDER_QTY  "
                + "FROM (SELECT pri.* FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE pri.PR_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }

        PLAN_PURCHASE_REQUISITION PRform = null;
        //取得新增的領料單號
        public PLAN_PURCHASE_REQUISITION getNewDeliveryOrderId(string prid, DateTime createDate)
        {
            using (var context = new topmepEntities())
            {
                PRform = context.PLAN_PURCHASE_REQUISITION.SqlQuery("SELECT * FROM PLAN_PURCHASE_REQUISITION pr WHERE pr.PARENT_PR_ID =@prid " +
                    "AND pr.PR_ID LIKE 'DF%' AND pr.CREATE_DATE = @createDate "
                   , new SqlParameter("prid", prid), new SqlParameter("createDate", createDate)).FirstOrDefault();
            }
            return PRform;
        }

        //取得領料單資料
        public List<PRFunction> getDOByPrjId(string projectid, string recipient, string prid, string caution)
        {

            logger.Info("search delivery form by 領料說明 =" + caution + ", 領料單編號 =" + prid + ", 領料人所屬單位 =" + recipient);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, RECIPIENT, CAUTION, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE PROJECT_ID =@projectid AND PR_ID LIKE 'D%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //領料說明查詢條件
            if (null != caution && caution != "")
            {
                sql = sql + "AND CAUTION LIKE @caution ";
                parameters.Add(new SqlParameter("caution", '%' + caution + '%'));
            }
            //領料單編號條件
            if (null != prid && prid != "")
            {
                sql = sql + "AND PR_ID =@prid ";
                parameters.Add(new SqlParameter("prid", prid));
            }
            //領料人條件
            if (null != recipient && recipient != "")
            {
                sql = sql + "AND RECIPIENT LIKE @recipient ";
                parameters.Add(new SqlParameter("recipient", '%' + recipient + '%'));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get delivery form sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get delivery form count=" + lstForm.Count);
            return lstForm;
        }

        //取得個別物料的庫存數量
        public PurchaseRequisition getInventoryByItemId(string itemid)
        {

            logger.Info("search item inventory by planitemid  =" + itemid);
            PurchaseRequisition lstItem = new PurchaseRequisition();
            //處理SQL 預先填入專案代號,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PurchaseRequisition>("SELECT pri.PLAN_ITEM_ID, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN, SUM(pri.RECEIPT_QTY) - ISNULL(A.DELIVERY_QTY, 0) AS INVENTORY_QTY, SUM(pri.RECEIPT_QTY) AS ALL_RECEIPT_QTY " +
                "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                "LEFT JOIN (SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid " +
                "GROUP BY pid.PLAN_ITEM_ID)A ON pri.PLAN_ITEM_ID = A.PLAN_ITEM_ID GROUP BY pri.PLAN_ITEM_ID, A.DELIVERY_QTY, " +
                "pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN HAVING pri.PLAN_ITEM_ID =@itemid; "
            , new SqlParameter("itemid", itemid)).First();
            }

            return lstItem;
        }

        //取得3天內須驗收的物料品項
        public List<PurchaseRequisition>getPlanItemByNeedDate(string prjid)
        {

            logger.Info(" get materials ready to receive in 3 days by 專案編號 =" + prjid);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pri.NEED_QTY, CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, pri.REMARK, pri.PR_ID, " +
                "pi.PLAN_ITEM_ID, pi.ITEM_DESC, pi.ITEM_ID, pi.ITEM_UNIT FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi " +
                "ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID WHERE pri.PR_ID LIKE 'PR%' AND pi.PROJECT_ID =@prjid " +
                "AND pri.NEED_DATE BETWEEN GETDATE() AND GETDATE()-3 ORDER BY pri.PR_ID ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prjid", prjid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get materials ready to receive in 3 days sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get material's item count ready to receive in 3 days =" + lstItem.Count);
            return lstItem;
        }

        #endregion

        #region 估驗

        public PLAN_ESTIMATION_FORM formEST = null;
        public List<EstimationForm> ESTItem = null;

        //取得個別廠商合約內容(含工資)
        public List<plansummary> getAllPlanContract(string projectid, string formName, string supplier)
        {
            logger.Info("search contract by 採購項目 =" + formName + "search contract by 供應商 =" + supplier);
            List<plansummary> lst = new List<plansummary>();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT  ' ' + p.PROJECT_ID + p.SUPPLIER_ID + p.FORM_NAME AS CONTRACT_ID, p.SUPPLIER_ID, p.FORM_NAME, p.SUPPLIER_ID + '_' + p.FORM_NAME AS CONTRACT_NAME, '' AS TYPE," +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY p.SUPPLIER_ID) AS NO FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid " +
                    "GROUP BY p.PROJECT_ID, p.SUPPLIER_ID, p.FORM_NAME HAVING p.SUPPLIER_ID IS NOT NULL ";

            //供應商
            if (null != supplier && supplier != "")
            {
                sql = sql + "AND p.SUPPLIER_ID LIKE @supplier ";
                parameters.Add(new SqlParameter("supplier", "%" + supplier + "%"));
            }
            //採購項目
            if (null != formName && formName != "")
            {
                sql = sql + "AND p.FORM_NAME LIKE @formName ";
                parameters.Add(new SqlParameter("formName", "%" + formName + "%"));
            }
            sql = sql + "UNION SELECT  'W' + p.PROJECT_ID + p.MAN_SUPPLIER_ID + p.MAN_FORM_NAME AS CONTRACT_ID, p.MAN_SUPPLIER_ID, p.MAN_FORM_NAME, p.MAN_SUPPLIER_ID + '_' + p.MAN_FORM_NAME AS CONTRACT_NAME, '工資' AS TYPE, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY p.MAN_SUPPLIER_ID) AS NO FROM PLAN_ITEM p WHERE p.PROJECT_ID =@projectid and p.MAN_PRICE IS NOT NULL " +
                    "GROUP BY p.PROJECT_ID, p.MAN_SUPPLIER_ID, p.MAN_FORM_NAME HAVING p.MAN_SUPPLIER_ID IS NOT NULL ";
            //供應商
            if (null != supplier && supplier != "")
            {
                sql = sql + "AND p.MAN_SUPPLIER_ID LIKE @supplierForWage ";
                parameters.Add(new SqlParameter("supplierForWage", "%" + supplier + "%"));
            }
            //採購項目
            if (null != formName && formName != "")
            {
                sql = sql + "AND p.MAN_FORM_NAME LIKE @formNameForWage ";
                parameters.Add(new SqlParameter("formNameForWage", "%" + formName + "%"));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get contract sql=" + sql);
                lst = context.Database.SqlQuery<plansummary>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get contract count=" + lst.Count);
            return lst;
        }

        //取得個別合約的明細資料
        public List<EstimationForm> getContractItemById(string contractid)
        {

            logger.Info("get contract item by contractid  =" + contractid);
            List<EstimationForm> lstItem = new List<EstimationForm>();
            //處理SQL 預先填入合約代號,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<EstimationForm>("SELECT pi.*, map.QTY AS mapQty, A.CUM_QTY AS CUM_EST_QTY, B.CUM_QTY AS CUM_RECPT_QTY, ISNULL(B.CUM_QTY, 0)-ISNULL(A.CUM_QTY,0) AS Quota FROM PLAN_ITEM pi " +
                    "LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN (SELECT ei.PLAN_ITEM_ID, SUM(ei.EST_QTY) AS CUM_QTY " +
                    "FROM PLAN_ESTIMATION_ITEM ei LEFT JOIN PLAN_ESTIMATION_FORM ef ON ei.EST_FORM_ID = ef.EST_FORM_ID JOIN TND_SUPPLIER sup ON SUBSTRING(ef.CONTRACT_ID, 7, 7) = sup.SUPPLIER_ID " +
                    "WHERE STUFF(ef.CONTRACT_ID, 7, 7, sup.COMPANY_NAME) = @contractid GROUP BY ei.PLAN_ITEM_ID)A ON pi.PLAN_ITEM_ID = A.PLAN_ITEM_ID " +
                    "LEFT JOIN (SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS CUM_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr " +
                    "ON pri.PR_ID = pr.PR_ID WHERE pri.PR_ID LIKE 'RP%' AND pr.PROJECT_ID = SUBSTRING(@contractid,1,6) GROUP BY pri.PLAN_ITEM_ID)B ON pi.PLAN_ITEM_ID = B.PLAN_ITEM_ID WHERE " +
                    "pi.PROJECT_ID + pi.SUPPLIER_ID + pi.FORM_NAME = @contractid OR pi.PROJECT_ID + pi.MAN_SUPPLIER_ID + pi.MAN_FORM_NAME = @contractid ; "
            , new SqlParameter("contractid", contractid)).ToList();
            }

            return lstItem;
        }

        //取得個別材料廠商合約資料與金額
        public plansummary getPlanContract4Est(string contractid)
        {
            plansummary lst = new plansummary();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT  A.PROJECT_ID + A.ID + A.FORM_NAME AS CONTRACT_ID, A.SUPPLIER_ID, A.FORM_NAME, " +
                    "SUM(A.mapQty * A.ITEM_UNIT_COST) MATERIAL_COST, SUM(A.mapQty * ISNULL(A.MAN_PRICE, 0)) WAGE_COST, " +
                    "SUM(A.ITEM_QUANTITY * A.ITEM_UNIT_PRICE) REVENUE, SUM(A.mapQty * A.ITEM_UNIT_COST * A.BUDGET_RATIO / 100) BUDGET, " +
                    "(SUM(A.mapQty * A.ITEM_UNIT_COST) + SUM(A.mapQty * ISNULL(A.MAN_PRICE, 0))) COST, (SUM(A.ITEM_QUANTITY * A.ITEM_UNIT_PRICE) - " +
                    "SUM(A.mapQty * A.ITEM_UNIT_COST) - SUM(A.mapQty * ISNULL(A.MAN_PRICE, 0))) PROFIT, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY A.SUPPLIER_ID) AS NO FROM (SELECT pi.*, s.SUPPLIER_ID AS ID, map.QTY AS mapQty FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                    "pi.SUPPLIER_ID = s.COMPANY_NAME LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID)A GROUP BY A.PROJECT_ID, A.ID, A.FORM_NAME, A.SUPPLIER_ID HAVING A.PROJECT_ID + A.SUPPLIER_ID + A.FORM_NAME =@contractid ; "
                   , new SqlParameter("contractid", contractid)).First();
            }
            return lst;
        }
        //取得個別工資廠商合約資料與金額
        public plansummary getPlanContractOfWage4Est(string contractid)
        {
            plansummary lst = new plansummary();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT  A.PROJECT_ID + A.ID + A.MAN_FORM_NAME AS CONTRACT_ID, A.MAN_SUPPLIER_ID, A.MAN_FORM_NAME, " +
                    "SUM(A.mapQty * ISNULL(A.MAN_PRICE, 0)) WAGE_COST, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY A.MAN_SUPPLIER_ID) AS NO FROM(SELECT pi.*, s.SUPPLIER_ID AS ID, map.QTY AS mapQty FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                    "pi.MAN_SUPPLIER_ID = s.COMPANY_NAME LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID)A GROUP BY A.PROJECT_ID, A.MAN_SUPPLIER_ID, A.MAN_FORM_NAME, A.ID HAVING A.PROJECT_ID + A.MAN_SUPPLIER_ID + A.MAN_FORM_NAME =@contractid ; "
                   , new SqlParameter("contractid", contractid)).First();
            }
            return lst;
        }
        string sno_key = "EST";
        public string getEstNo()
        {
            string estNo = null;
            //取得估驗單編號
            using (var context = new topmepEntities())
            {
                SerialKeyService snoservice = new SerialKeyService();
                estNo = snoservice.getSerialKey(sno_key);
            }
            return estNo;
        }

        // 寫入估驗內容
        public string newEST(string formid, PLAN_ESTIMATION_FORM form, string[] lstItemId)
        {
            //1.建立估驗單
            logger.Info("create new estimation form ");
            using (var context = new topmepEntities())
            {
                context.PLAN_ESTIMATION_FORM.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase Requisition=" + i);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_ESTIMATION_ITEM> lstItem = new List<PLAN_ESTIMATION_ITEM>();
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

                string sql = "INSERT INTO PLAN_ESTIMATION_ITEM (EST_FORM_ID, PLAN_ITEM_ID) "
                + "SELECT '" + formid + "' as EST_FORM_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID  "
                + "FROM (SELECT pi.PLAN_ITEM_ID FROM PLAN_ITEM pi WHERE pi.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return formid;
            }
        }

        //更新估驗數量
        public int refreshEST(string formid, PLAN_ESTIMATION_FORM form, List<PLAN_ESTIMATION_ITEM> lstItem)
        {
            logger.Info("Update plan estimation form id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan estimation form =" + i);
                    logger.Info("purchase estimation item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_ESTIMATION_ITEM item in lstItem)
                    {
                        PLAN_ESTIMATION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_ESTIMATION_ITEM excelItem = context.PLAN_ESTIMATION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ESTIMATION_ITEM.Find(excelItem.EST_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.EST_QTY = item.EST_QTY;
                        context.PLAN_ESTIMATION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update plan estimation item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new estimation form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //取得符合條件之估驗單名單
        public List<ESTFunction> getESTListByEstId(string projectid, string contractid, string estid, int status)
        {
            logger.Info("search estimation form by 估驗單編號 =" + estid + ", 合約名稱 =" + contractid + ", 估驗單狀態 =" + status);
            List<ESTFunction> lstForm = new List<ESTFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            if (20 == status)
            {
                string sql = "SELECT CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.EST_FORM_ID, A.STATUS, A.CONTRACT_NAME, A.SUPPLIER_NAME, ROW_NUMBER() OVER(ORDER BY A.EST_FORM_ID) AS NO " +
                    "FROM (SELECT ef.CREATE_DATE, ef.EST_FORM_ID, ef.STATUS, STUFF(ef.CONTRACT_ID,7, 7, sup.COMPANY_NAME) AS CONTRACT_NAME, sup.COMPANY_NAME AS SUPPLIER_NAME " +
                    "FROM PLAN_ESTIMATION_FORM ef LEFT JOIN TND_SUPPLIER sup ON SUBSTRING(ef.CONTRACT_ID, 7, 7) = sup.SUPPLIER_ID WHERE ef.PROJECT_ID =@projectid)A ";


                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                sql = sql + "WHERE A.STATUS > 10 ";

                //估驗單編號條件
                if (null != estid && estid != "")
                {
                    sql = sql + "AND A.EST_FORM_ID =@estid ";
                    parameters.Add(new SqlParameter("estid", estid));
                }
                //合約名稱條件
                if (null != contractid && contractid != "")
                {
                    sql = sql + "AND A.CONTRACT_NAME LIKE @contractid ";
                    parameters.Add(new SqlParameter("contractid", '%' + contractid + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get estimation form sql=" + sql);
                    lstForm = context.Database.SqlQuery<ESTFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get estimation form count=" + lstForm.Count);
            }
            else
            {
                string sql = "SELECT CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.EST_FORM_ID, A.STATUS, A.CONTRACT_NAME, A.SUPPLIER_NAME, ROW_NUMBER() OVER(ORDER BY A.EST_FORM_ID) AS NO " +
                    "FROM (SELECT ef.CREATE_DATE, ef.EST_FORM_ID, ef.STATUS, STUFF(ef.CONTRACT_ID,7, 7, sup.COMPANY_NAME) AS CONTRACT_NAME, sup.COMPANY_NAME AS SUPPLIER_NAME " +
                    "FROM PLAN_ESTIMATION_FORM ef LEFT JOIN TND_SUPPLIER sup ON SUBSTRING(ef.CONTRACT_ID, 7, 7) = sup.SUPPLIER_ID WHERE ef.PROJECT_ID =@projectid)A ";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                sql = sql + "WHERE A.STATUS < 20 ";

                using (var context = new topmepEntities())
                {
                    logger.Debug("get estimation form sql=" + sql);
                    lstForm = context.Database.SqlQuery<ESTFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get estimation form count=" + lstForm.Count);
            }
            return lstForm;
        }

        //取得估驗單資料
        public void getESTByEstId(string estid)
        {
            logger.Info("get form : formid=" + estid);
            formEST = null;
            using (var context = new topmepEntities())
            {
                //取得估驗單檔頭資訊
                string sql = "SELECT EST_FORM_ID, PROJECT_ID, CONTRACT_ID, PLUS_TAX, TAX_AMOUNT, TAX_RATIO, PAYMENT_TRANSFER, FOREIGN_PAYMENT, RETENTION_PAYMENT, MODIFY_DATE, REMARK, INVOICE, " +
                    "CREATE_ID, CREATE_DATE, SETTLEMENT, STATUS, TYPE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID =@estid ";

                formEST = context.PLAN_ESTIMATION_FORM.SqlQuery(sql, new SqlParameter("estid", estid)).FirstOrDefault();
                //取得估驗單明細
                ESTItem = context.Database.SqlQuery<EstimationForm>("SELECT pei.PLAN_ITEM_ID, pei.EST_QTY, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.ITEM_FORM_QUANTITY, " +
                    "pi.ITEM_UNIT_PRICE, A.CUM_QTY AS CUM_QTY FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ESTIMATION_FORM ef ON pei.EST_FORM_ID = ef.EST_FORM_ID LEFT JOIN PLAN_ITEM pi ON " +
                    "pei.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN (SELECT pei.PLAN_ITEM_ID, sum(pei.EST_QTY) AS CUM_QTY FROM PLAN_ESTIMATION_ITEM pei JOIN PLAN_ESTIMATION_FORM ef ON " +
                    "pei.EST_FORM_ID = ef.EST_FORM_ID WHERE ef.CREATE_DATE < (select CREATE_DATE from PLAN_ESTIMATION_FORM where EST_FORM_ID = @estid) GROUP BY  pei.PLAN_ITEM_ID)A " +
                    "ON pei.PLAN_ITEM_ID = A.PLAN_ITEM_ID WHERE pei.EST_FORM_ID = @estid", new SqlParameter("estid", estid)).ToList();
                logger.Debug("get estimation form item count:" + ESTItem.Count);
            }
        }

        public int addOtherPayment(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增其他扣款資料
            int i = 0;
            logger.Info("add other payment = " + lstItem.Count);
            //2.將扣款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.TYPE = "O";
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add other payment count =" + i);
            return i;
        }
        //取得估驗單其他扣款明細資料
        public List<PLAN_OTHER_PAYMENT> getOtherPayById(string id)
        {

            logger.Info("get other payment by EST id + contractid  =" + id);
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_OTHER_PAYMENT>("SELECT * FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID =@id AND TYPE = 'O' ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //取得估驗單狀態
        public int getStatusById(string id)
        {
            int status = -10;
            logger.Info("get EST status by EST id + contractid  =" + id);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                status = context.Database.SqlQuery<int>("SELECT STATUS FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID + CONTRACT_ID =@id ; "
            , new SqlParameter("id", id)).FirstOrDefault();
            }

            return status;
        }
        public int delOtherPayByESTId(string estid)
        {
            logger.Info("remove all other payment detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these other payment record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID=@estid AND TYPE = 'O' ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN OTHER PAYMENT count=" + i);
            return i;
        }

        //更新估驗數量
        public int refreshESTQty(string formid, List<PLAN_ESTIMATION_ITEM> lstItem)
        {
            logger.Info("Update estiomation items, it's est form id =" + formid);
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    //將item資料寫入 
                    foreach (PLAN_ESTIMATION_ITEM item in lstItem)
                    {
                        PLAN_ESTIMATION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_ESTIMATION_ITEM excelItem = context.PLAN_ESTIMATION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ESTIMATION_ITEM.Find(excelItem.EST_ITEM_ID);

                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        if (item.EST_QTY != null)
                        {
                            existItem.EST_QTY = item.EST_QTY;
                        }
                        context.PLAN_ESTIMATION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update estimation item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new est form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return j;
        }

        //取得新估驗單估驗次數
        int est = 0;
        public int getEstCountById(string contractid)
        {
            using (var context = new topmepEntities())
            {
                est = context.Database.SqlQuery<int>("SELECT ISNULL((SELECT COUNT(CONTRACT_ID) FROM PLAN_ESTIMATION_FORM " +
                    "GROUP BY CONTRACT_ID HAVING CONTRACT_ID =@contractid),0)+1 AS EST_COUNT "
                   , new SqlParameter("contractid", contractid)).First();
            }
            return est;
        }

        //取得現有估驗單之估驗次數
        int estcount = 0;
        public int getEstCountByESTId(string estid)
        {
            using (var context = new topmepEntities())
            {
                estcount = context.Database.SqlQuery<int>("SELECT ISNULL((SELECT COUNT(ef.CONTRACT_ID) + 1 FROM PLAN_ESTIMATION_FORM ef WHERE ef.CREATE_DATE < " +
                    "(select CREATE_DATE from PLAN_ESTIMATION_FORM where EST_FORM_ID =@estid) AND ef.CONTRACT_ID = (select CONTRACT_ID from PLAN_ESTIMATION_FORM " +
                    "where EST_FORM_ID =@estid) GROUP BY ef.CONTRACT_ID),1) AS EST_COUNT  "
                   , new SqlParameter("estid", estid)).First();
            }
            return estcount;
        }

        AdvancePaymentFunction advancePay = null;
        //取得估驗單預付款明細資料
        public AdvancePaymentFunction getAdvancePayById(string id)
        {

            logger.Info("get advance payment by EST id + contractid  =" + id);
            using (var context = new topmepEntities())
            {
                advancePay = context.Database.SqlQuery<AdvancePaymentFunction>("SELECT (SELECT AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID =@id AND TYPE = 'A') AS A_AMOUNT, " +
                    "(SELECT AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID =@id AND TYPE = 'B') AS B_AMOUNT, " +
                    "(SELECT AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID =@id AND TYPE = 'C') AS C_AMOUNT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = SUBSTRING(@id,10, LEN(@id)-9) AND TYPE = 'A' " +
                    "AND CREATE_DATE < ISNULL((SELECT CREATE_DATE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = SUBSTRING(@id, 1,9) AND TYPE = 'A'), GETDATE()) GROUP BY CONTRACT_ID),0) CUM_A_AMOUNT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = SUBSTRING(@id,10, LEN(@id)-9) AND TYPE = 'B' " +
                    "AND CREATE_DATE < ISNULL((SELECT CREATE_DATE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = SUBSTRING(@id, 1,9) AND TYPE = 'B'), GETDATE()) GROUP BY CONTRACT_ID),0) CUM_B_AMOUNT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = SUBSTRING(@id,10, LEN(@id)-9) AND TYPE = 'C' " +
                    "AND CREATE_DATE < ISNULL((SELECT CREATE_DATE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = SUBSTRING(@id, 1,9) AND TYPE = 'C'), GETDATE()) GROUP BY CONTRACT_ID),0) CUM_C_AMOUNT  "
            , new SqlParameter("id", id)).First();
            }

            return advancePay;
        }

        public int addAdvancePayment(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增預付款資料
            int i = 0;
            logger.Info("add advance payment = " + lstItem.Count);
            //2.將預付款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add advance payment count =" + i);
            return i;
        }

        //取得估驗單是否有預付款資料
        public List<PLAN_OTHER_PAYMENT> getAdvancePayByESTId(string id)
        {

            logger.Info("get advance payment by EST id + contractid  =" + id);
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_OTHER_PAYMENT>("SELECT * FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID =@id AND TYPE IN ('A','B','C') ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //更新預付款資料
        public int updateAdvancePayment(string estid, List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.修改預付款資料
            int i = 0;
            logger.Info("update advance payment = " + lstItem.Count);
            //2.將預付款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    PLAN_OTHER_PAYMENT existItem = null;
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("estid", estid));
                    parameters.Add(new SqlParameter("type", item.TYPE));
                    string sql = "SELECT * FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @estid and TYPE = @type ";
                    logger.Info(sql + " ;" + item.EST_FORM_ID + item.TYPE);
                    PLAN_OTHER_PAYMENT excelItem = context.PLAN_OTHER_PAYMENT.SqlQuery(sql, parameters.ToArray()).First();
                    existItem = context.PLAN_OTHER_PAYMENT.Find(excelItem.OTHER_PAYMENT_ID);
                    logger.Debug("find exist item=" + existItem.TYPE);
                    existItem.AMOUNT = item.AMOUNT;
                    context.PLAN_OTHER_PAYMENT.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update advance payment count =" + i);
            return i;
        }

        PaymentDetailsFunction detailsPay = null;
        //取得估驗單預付款明細資料
        public PaymentDetailsFunction getDetailsPayById(string id)
        {
            logger.Info("get details payment by  id  =" + id);
            using (var context = new topmepEntities())
            {
                detailsPay = context.Database.SqlQuery<PaymentDetailsFunction>("SELECT D.*,(D.EST_AMOUNT-D.T_FOREIGN+D.T_REPAYMENT) AS SUB_AMOUNT, (D.CUM_EST_AMOUNT - D.CUM_T_FOREIGN + D.CUM_T_REPAYMENT) AS CUM_SUB_AMOUNT, (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER+D.T_REPAYMENT-D.T_REFUND) AS PAYABLE_AMOUNT, " +
                    "(D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN-D.CUM_T_RETENTION-D.CUM_T_ADVANCE-D.CUM_T_OTHER+D.CUM_T_REPAYMENT-D.CUM_T_REFUND) AS CUM_PAYABLE_AMOUNT, " +
                    "(D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER + D.TAX_AMOUNT+D.T_REPAYMENT-D.T_REFUND) AS PAID_AMOUNT, (D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN-D.CUM_T_RETENTION-D.CUM_T_ADVANCE-D.CUM_T_OTHER + D.CUM_TAX_AMOUNT+D.CUM_T_REPAYMENT-D.CUM_T_REFUND) AS CUM_PAID_AMOUNT, " +
                    "(D.T_FOREIGN + D.CUM_T_FOREIGN) AS TOTAL_FOREIGN, (D.EST_AMOUNT - D.T_FOREIGN + D.CUM_EST_AMOUNT - D.CUM_T_FOREIGN + D.T_REPAYMENT + D.CUM_T_REPAYMENT) AS TOTAL_SUB_AMOUNT, (D.T_RETENTION + D.CUM_T_RETENTION) AS TOTAL_RETENTION, (D.T_ADVANCE + D.CUM_T_ADVANCE) AS TOTAL_ADVANCE, (D.T_REPAYMENT + D.CUM_T_REPAYMENT) AS TOTAL_REPAYMENT, " +
                    "(D.T_REFUND + D.CUM_T_REFUND) AS TOTAL_REFUND, (D.T_OTHER + D.CUM_T_OTHER) AS TOTAL_OTHER, (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER + D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN - D.CUM_T_RETENTION - D.CUM_T_ADVANCE - D.CUM_T_OTHER + D.CUM_T_REPAYMENT - D.CUM_T_REFUND) AS TOTAL_PAYABLE_AMOUNT,  " +
                    "(D.TAX_AMOUNT + D.CUM_TAX_AMOUNT) AS TOTAL_TAX_AMOUNT, (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER + D.TAX_AMOUNT + D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN-D.CUM_T_RETENTION-D.CUM_T_ADVANCE-D.CUM_T_OTHER + D.CUM_TAX_AMOUNT + D.CUM_T_REPAYMENT - D.CUM_T_REFUND) AS TOTAL_PAID_AMOUNT " +
                    "FROM(SELECT ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID = @id AND TYPE IN('A', 'B', 'C') GROUP BY EST_FORM_ID + CONTRACT_ID), 0) AS T_ADVANCE, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) AND TYPE IN('A', 'B', 'C') " +
                    "AND CREATE_DATE < (SELECT ISNULL((SELECT MIN(CREATE_DATE) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9) AND TYPE IN('A', 'B', 'C')), GETDATE())) GROUP BY CONTRACT_ID),0) AS CUM_T_ADVANCE, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID = @id AND TYPE = 'R' GROUP BY EST_FORM_ID + CONTRACT_ID), 0) AS T_REPAYMENT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) AND TYPE = 'R' " +
                    "AND CREATE_DATE < (SELECT ISNULL((SELECT MIN(CREATE_DATE) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9) AND TYPE = 'R'), GETDATE())) GROUP BY CONTRACT_ID),0) AS CUM_T_REPAYMENT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID + CONTRACT_ID = @id AND TYPE = 'F' GROUP BY EST_FORM_ID + CONTRACT_ID), 0) AS T_REFUND, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) AND TYPE = 'F' " +
                    "AND CREATE_DATE < (SELECT ISNULL((SELECT MIN(CREATE_DATE) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9) AND TYPE = 'F'), GETDATE())) GROUP BY CONTRACT_ID),0) AS CUM_T_REFUND, " +
                    "ISNULL((SELECT SUM(AMOUNT) AS AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID IN(SELECT EST_FORM_ID FROM PLAN_ESTIMATION_FORM WHERE CREATE_DATE < (SELECT DISTINCT ef.CREATE_DATE FROM PLAN_OTHER_PAYMENT pop JOIN PLAN_ESTIMATION_FORM ef " +
                    "ON pop.EST_FORM_ID = ef.EST_FORM_ID WHERE pop.EST_FORM_ID = SUBSTRING(@id, 1, 9))) AND TYPE = 'O' AND CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9)),0) AS CUM_T_OTHER, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE TYPE = 'O' GROUP BY EST_FORM_ID HAVING EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS T_OTHER, " +
                    "ISNULL((SELECT TAX_RATIO FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS TAX_RATIO, " +
                    "ISNULL((SELECT RETENTION_PAYMENT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS T_RETENTION, " +
                    "ISNULL((SELECT TAX_AMOUNT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS TAX_AMOUNT, " +
                    "ISNULL((SELECT SUM(RETENTION_PAYMENT) FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9))),0)  AS CUM_T_RETENTION, " +
                    "ISNULL((SELECT SUM(TAX_AMOUNT) FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9))),0) AS CUM_TAX_AMOUNT, " +
                    "ISNULL((SELECT SUM(pei.EST_QTY * pi.ITEM_UNIT_COST) PRICE FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ITEM pi ON PEI.PLAN_ITEM_ID = PI.PLAN_ITEM_ID WHERE pei.EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS EST_AMOUNT, " +
                    "ISNULL((SELECT SUM(pei.EST_QTY * pi.ITEM_UNIT_COST) PRICE FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ITEM pi ON pei.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN PLAN_ESTIMATION_FORM ef ON pei.EST_FORM_ID = ef.EST_FORM_ID WHERE ef.CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9))), 0) AS CUM_EST_AMOUNT,  " +
                    "ISNULL((SELECT FOREIGN_PAYMENT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS T_FOREIGN, ISNULL((SELECT SUM(FOREIGN_PAYMENT) FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) " +
                    "AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9))),0) AS CUM_T_FOREIGN)D "
            , new SqlParameter("id", id)).First();
            }
            return detailsPay;
        }

        decimal retention = 0;
        public decimal getRetentionAmountById(string id)
        {
            using (var context = new topmepEntities())
            {
                retention = context.Database.SqlQuery<decimal>("SELECT RATIO * AMOUNT / 100 * 1.05 FROM(SELECT ISNULL((SELECT PAYMENT_RETENTION_RATIO FROM PLAN_PAYMENT_TERMS WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9)), 0) AS RATIO, " +
                    "(SELECT ISNULL(SUM(pei.EST_QTY * pi.ITEM_UNIT_COST),0) PRICE FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ITEM pi ON pei.PLAN_ITEM_ID = pi.PLAN_ITEM_ID WHERE pei.EST_FORM_ID = SUBSTRING(@id, 1, 9)) AS AMOUNT)B  "
                   , new SqlParameter("id", id)).FirstOrDefault();
            }
            return retention;
        }
        //寫入估驗保留款
        public int UpdateRetentionAmountById(string id)
        {
            int i = 0;
            logger.Info("update retention payment of EST form by id" + id);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET RETENTION_PAYMENT = i.PAY, TAX_AMOUNT = i.TAX_AMOUNT  " +
                "FROM(SELECT RATIO * AMOUNT / 100 * (1+TAX_RATIO/100) AS PAY, (AMOUNT -T_FOREIGN) * TAX_RATIO/100 AS TAX_AMOUNT FROM(SELECT ISNULL((SELECT PAYMENT_RETENTION_RATIO FROM PLAN_PAYMENT_TERMS WHERE " +
                "CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9)), 0) AS RATIO, ISNULL((SELECT TAX_RATIO FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS TAX_RATIO, " +
                "ISNULL((SELECT FOREIGN_PAYMENT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)),0) AS T_FOREIGN, " +
                "(SELECT ISNULL(SUM(pei.EST_QTY * pi.ITEM_UNIT_COST), 0) PRICE FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ITEM pi ON pei.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                "WHERE pei.EST_FORM_ID = SUBSTRING(@id, 1, 9)) AS AMOUNT)B) i WHERE EST_FORM_ID = SUBSTRING(@id, 1, 9)";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("id", id));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        public int delESTByESTId(string estid)
        {
            logger.Info("remove EST form detail by EST FORM ID =" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete EST form record by est form id =" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @estid ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN ESTIMATION FORM count=" + i);
            return i;
        }

        public int delESTItemsByESTId(string estid)
        {
            logger.Info("remove EST items detail by EST FORM ID  =" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete EST items record by est form id =" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID = @estid ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN ESTIMATION ITEM count=" + i);
            return i;
        }
        //更新估驗單狀態為草稿
        public int UpdateESTStatusById(string estid)
        {
            int i = 0;
            logger.Info("update the status of EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 10 WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新估驗單狀態為送審
        public int RefreshESTStatusById(string estid)
        {
            int i = 0;
            logger.Info("update the status of EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 20 WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //修改估驗單內容
        public int RefreshESTByEstId(string estid, string tax, decimal taxratio)
        {
            int i = 0;
            logger.Info("update EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET PLUS_TAX = @tax, TAX_RATIO =@taxratio WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            parameters.Add(new SqlParameter("tax", tax));
            parameters.Add(new SqlParameter("taxratio", taxratio));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //修改估驗單額外扣款
        public int RefreshESTAmountByEstId(string estid, decimal foreign_payment, decimal retention, decimal tax_amount, string remark)
        {
            int i = 0;
            logger.Info("update EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET FOREIGN_PAYMENT = @foreign_payment, RETENTION_PAYMENT = @retention, TAX_AMOUNT =@tax_amount, REMARK =@remark WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            parameters.Add(new SqlParameter("foreign_payment", foreign_payment));
            parameters.Add(new SqlParameter("retention", retention));
            parameters.Add(new SqlParameter("tax_amount", tax_amount));
            parameters.Add(new SqlParameter("remark", remark));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新估驗單狀態為退件
        public int RejectESTByEstId(string estid)
        {
            int i = 0;
            logger.Info("reject EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 0 WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新估驗單狀態為已核可
        public int ApproveESTByEstId(string estid)
        {
            int i = 0;
            logger.Info("Approve EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 30, MODIFY_DATE = GETDATE() WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }
        //取得估驗單憑證資料
        public List<PLAN_INVOICE> getInvoiceById(string id)
        {

            logger.Info("get invoice by EST id + contractid  =" + id);
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_INVOICE>("SELECT INVOICE_ID, EST_FORM_ID, CONTRACT_ID, INVOICE_DATE, INVOICE_NUMBER, " +
                    "AMOUNT, TAX, TYPE, SUB_TYPE, PLAN_ITEM_ID, DISCOUNT_QTY, DISCOUNT_UNIT_PRICE, CREATE_DATE FROM PLAN_INVOICE WHERE EST_FORM_ID + CONTRACT_ID =@id ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        public int addInvoice(List<PLAN_INVOICE> lstItem)
        {
            //1.新增憑證資料
            int i = 0;
            logger.Info("add invoice = " + lstItem.Count);
            //2.將扣款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_INVOICE item in lstItem)
                {
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_INVOICE.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add invoice count =" + i);
            return i;
        }

        //取得付款條件
        public string getTermsByContractId(string contractid)
        {
            string terms = null;
            logger.Info("get payment terms by contractid  =" + contractid);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                terms = context.Database.SqlQuery<string>("SELECT PAYMENT_TERMS FROM PLAN_PAYMENT_TERMS WHERE CONTRACT_ID =@contractid ; "
            , new SqlParameter("contractid", contractid)).FirstOrDefault();
            }

            return terms;
        }

        //取得估驗單代付支出明細資料
        public List<RePaymentFunction> getRePaymentById(string id)
        {

            logger.Info("get repayment by EST id + contractid  =" + id);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID_FOR_REFUND AS CONTRACT_ID_FOR_REFUND, s.COMPANY_NAME AS COMPANY_NAME " +
                    "FROM PLAN_OTHER_PAYMENT pop LEFT JOIN TND_SUPPLIER s ON SUBSTRING(pop.CONTRACT_ID_FOR_REFUND, 7, 7) = s.SUPPLIER_ID WHERE pop.EST_FORM_ID + pop.CONTRACT_ID =@id AND pop.TYPE = 'R' ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //取得專案發包廠商資料
        public List<RePaymentFunction> getSupplierOfContractByPrjId(string prjid)
        {

            logger.Info("get repayment by projectid  =" + prjid);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT DISTINCT pi.SUPPLIER_ID AS COMPANY_NAME, pi.PROJECT_ID + s.SUPPLIER_ID + REPLACE(pi.FORM_NAME,',','*') AS CONTRACT_NAME " +
                    "FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON pi.SUPPLIER_ID = s.COMPANY_NAME  WHERE PROJECT_ID =@prjid UNION SELECT DISTINCT pi.MAN_SUPPLIER_ID, " +
                    "pi.PROJECT_ID + s.SUPPLIER_ID + REPLACE(pi.MAN_FORM_NAME,',','*') FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON pi.MAN_SUPPLIER_ID = s.COMPANY_NAME WHERE PROJECT_ID =@prjid ; "
            , new SqlParameter("prjid", prjid)).ToList();
            }

            return lstItem;
        }

        public int AddRePay(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增代付支出資料
            int i = 0;
            logger.Info("add repayment = " + lstItem.Count);
            //2.將代付支出資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.TYPE = "R";
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add repayment count =" + i);
            return i;
        }

        public int delRePayByESTId(string estid)
        {
            logger.Info("remove all repayment detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these repayment record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID=@estid AND TYPE = 'R' ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN OTHER PAYMENT count=" + i);
            return i;
        }

        //取得估驗單代付扣回明細資料
        public List<RePaymentFunction> getRefundById(string id)
        {

            logger.Info("get refund by EST id + contractid  =" + id);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT A.*, ISNULL((SELECT COUNT(ef.CONTRACT_ID) + 1 FROM PLAN_ESTIMATION_FORM ef WHERE ef.CREATE_DATE <  " +
                    "(SELECT DISTINCT pef.CREATE_DATE from PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM pef ON pef.EST_FORM_ID = pop.EST_FORM_ID_REFUND WHERE pop.EST_FORM_ID = SUBSTRING(@id,1,9) AND pef.CREATE_DATE IS NOT NULL) " +
                    "AND ef.CONTRACT_ID = (SELECT DISTINCT pef.CONTRACT_ID FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM pef ON pef.EST_FORM_ID = pop.EST_FORM_ID_REFUND WHERE pop.EST_FORM_ID = SUBSTRING(@id,1,9) AND pef.CREATE_DATE IS NOT NULL) " +
                    "GROUP BY ef.CONTRACT_ID),1) AS EST_COUNT_REFUND FROM(SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID AS CONTRACT_ID, " +
                    "pop.EST_FORM_ID_REFUND AS EST_FORM_ID_REFUND, pop.CONTRACT_ID_FOR_REFUND AS CONTRACT_ID_FOR_REFUND, s.COMPANY_NAME AS COMPANY_NAME FROM PLAN_OTHER_PAYMENT pop LEFT JOIN TND_SUPPLIER s " +
                    "ON SUBSTRING(pop.CONTRACT_ID_FOR_REFUND, 7, 7) = s.SUPPLIER_ID " +
                    "WHERE pop.EST_FORM_ID + pop.CONTRACT_ID =@id AND pop.TYPE = 'F')A ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //取得特定廠商所有代付扣回明細資料
        public List<RePaymentFunction> getRefundOfSupplierById(string id)
        {

            logger.Info("get all refunds of this supplier by EST id + contractid  =" + id);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID AS CONTRACT_ID, " +
                    "pop.EST_FORM_ID_REFUND AS EST_FORM_ID_REFUND, pop.EST_COUNT_REFUND AS EST_COUNT_REFUND, s.COMPANY_NAME AS COMPANY_NAME " +
                    "FROM PLAN_OTHER_PAYMENT pop LEFT JOIN TND_SUPPLIER s ON SUBSTRING(pop.CONTRACT_ID_FOR_REFUND, 7, 7) = s.SUPPLIER_ID " +
                    "WHERE pop.CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9) AND pop.TYPE = 'F' AND pop.EST_COUNT_REFUND IS NOT NULL ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }
        //取得需扣回給代付廠商之資料
        public List<RePaymentFunction> getSupplierOfContractRefundById(string id)
        {

            logger.Info("get refund by contractid  =" + id);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID AS CONTRACT_ID, pop.EST_FORM_ID AS EST_FORM_ID_REFUND, " +
                    "pop.OTHER_PAYMENT_ID AS OTHER_PAYMENT_ID, s.COMPANY_NAME AS COMPANY_NAME FROM PLAN_OTHER_PAYMENT pop LEFT JOIN TND_SUPPLIER s " +
                    "ON SUBSTRING(pop.CONTRACT_ID, 7, 7) = s.SUPPLIER_ID " +
                    "WHERE REPLACE(pop.CONTRACT_ID_FOR_REFUND,'*',',') =@id AND pop.TYPE = 'R' AND pop.CONTRACT_ID + pop.EST_FORM_ID + REPLACE(pop.CONTRACT_ID_FOR_REFUND,'*',',') + pop.REASON NOT IN " +
                    "(SELECT REPLACE(op.CONTRACT_ID_FOR_REFUND,'*',',') + op.EST_FORM_ID_REFUND + op.CONTRACT_ID + op.REASON FROM PLAN_OTHER_PAYMENT op WHERE op.TYPE = 'F'); "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        public string AddRefund(string formid, string contractid, string[] lstItemId)
        {
            //寫入代付支出資料
            using (var context = new topmepEntities())
            {
                List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
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

                string sql = "INSERT INTO PLAN_OTHER_PAYMENT (EST_FORM_ID, CONTRACT_ID, AMOUNT, "
                    + "CONTRACT_ID_FOR_REFUND, REASON, EST_FORM_ID_REFUND, TYPE) "
                    + "SELECT '" + formid + "' AS EST_FORM_ID, '" + contractid + "' AS CONTRACT_ID, AMOUNT, CONTRACT_ID_FOR_REFUND, REASON, EST_FORM_ID AS EST_FORM_ID_REFUND, 'F' AS TYPE "
                    + "FROM PLAN_OTHER_PAYMENT where OTHER_PAYMENT_ID IN (" + ItemId + ")";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return formid;
            }
        }

        public int delRefundByESTId(string estid)
        {
            logger.Info("remove all refund detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these refund record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID=@estid AND TYPE = 'F' ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN OTHER PAYMENT count=" + i);
            return i;
        }

        public int RefreshRefund(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增代付扣回資料
            int i = 0;
            logger.Info("add refund = " + lstItem.Count);
            //2.將代付扣回資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.TYPE = "F";
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add refund count =" + i);
            return i;
        }

        //取得廠商合約需扣回之總金額
        public decimal getBalanceOfRefundById(string id)
        {

            logger.Info("get the balance of refund by contractid  =" + id);
            decimal balance = 0;
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                balance = context.Database.SqlQuery<decimal>("SELECT ISNULL(R_AMOUNT,0) - ISNULL(F_AMOUNT,0) AS BALANCE " +
                    "FROM (SELECT CONTRACT_ID_FOR_REFUND, SUM(AMOUNT) AS R_AMOUNT FROM PLAN_OTHER_PAYMENT WHERE TYPE = 'R' GROUP BY CONTRACT_ID_FOR_REFUND) B " +
                    "LEFT JOIN (SELECT CONTRACT_ID, SUM(AMOUNT) AS F_AMOUNT FROM PLAN_OTHER_PAYMENT WHERE TYPE = 'F' GROUP BY CONTRACT_ID)A " +
                    "ON REPLACE(B.CONTRACT_ID_FOR_REFUND,'*',',') = A.CONTRACT_ID  GROUP BY B.CONTRACT_ID_FOR_REFUND, R_AMOUNT, F_AMOUNT " +
                    "HAVING REPLACE(B.CONTRACT_ID_FOR_REFUND,'*',',') =@id AND ISNULL(R_AMOUNT,0) - ISNULL(F_AMOUNT,0) > 0 ; "
                   , new SqlParameter("id", id)).FirstOrDefault();
            }

            return balance;
        }

        //取得未核准之估驗單號碼
        public string getEstNoByContractId(string contractid)
        {
            string EstNo = null;
            logger.Info("get payment terms by contractid  =" + contractid);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                EstNo = context.Database.SqlQuery<string>("SELECT EST_FORM_ID FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID =@contractid AND STATUS < 30 ; "
            , new SqlParameter("contractid", contractid)).FirstOrDefault();
            }

            return EstNo;
        }

        public int delInvoiceByESTId(string estid)
        {
            logger.Info("remove all invoice detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these invoice record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_INVOICE WHERE EST_FORM_ID=@estid ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN INVOICE count=" + i);
            return i;
        }
        #endregion

        //取得6個月內現金流資料
        public List<CashFlowFunction> getCashFlow()
        {

            logger.Info("get cash flow order by date in six months !!");
            List<CashFlowFunction> lstItem = new List<CashFlowFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<CashFlowFunction>("SELECT C.DATE_CASHFLOW AS DATE_CASHFLOW, C.AMOUNT_INFLOW AS AMOUNT_INFLOW, C.AMOUNT_OUTFLOW AS AMOUNT_OUTFLOW, C.BALANCE AS BALANCE, C.RUNNING_TOTAL AS RUNNING_TOTAL " +
                    "FROM (SELECT CASHFLOW_1.DATE_CASHFLOW, CASHFLOW_1.AMOUNT_INFLOW, CASHFLOW_1.AMOUNT_OUTFLOW, CASHFLOW_1.BALANCE, SUM(CASHFLOW_2.BALANCE) RUNNING_TOTAL FROM (SELECT DISTINCT CONVERT(char(10), pa.PAYMENT_DATE, 111) AS DATE_CASHFLOW, " +
                    "ISNULL(A.AMOUNT_INFLOW,0) AS AMOUNT_INFLOW, ISNULL(B.AMOUNT_OUTFLOW, 0) AS AMOUNT_OUTFLOW, ISNULL(A.AMOUNT_INFLOW, 0) - ISNULL(B.AMOUNT_OUTFLOW, 0) AS BALANCE FROM PLAN_ACCOUNT pa LEFT JOIN " +
                    "(SELECT CONVERT(char(10), pla.PAYMENT_DATE, 111) AS DATE_INFLOW, SUM(pla.AMOUNT) AS AMOUNT_INFLOW FROM PLAN_ACCOUNT pla WHERE ISDEBIT = 'Y' AND STATUS <> 0 AND PAYMENT_DATE BETWEEN CONVERT(char(10), getdate(), 111) AND getdate() + 180 " +
                    "GROUP BY CONVERT(char(10), PAYMENT_DATE, 111))A ON CONVERT(char(10), pa.PAYMENT_DATE, 111) = A.DATE_INFLOW LEFT JOIN(SELECT CONVERT(char(10), pla.PAYMENT_DATE, 111) AS DATE_OUTFLOW, " +
                    "SUM(pla.AMOUNT) AS AMOUNT_OUTFLOW FROM PLAN_ACCOUNT pla WHERE ISDEBIT = 'N' AND STATUS <> 0 AND PAYMENT_DATE BETWEEN CONVERT(char(10), getdate(), 111) AND getdate() + 180 GROUP BY CONVERT(char(10), PAYMENT_DATE, 111))B " +
                    "ON CONVERT(char(10), pa.PAYMENT_DATE, 111) = B.DATE_OUTFLOW)CASHFLOW_1, (SELECT DISTINCT CONVERT(char(10), pa.PAYMENT_DATE, 111) AS DATE_CASHFLOW, ISNULL(A.AMOUNT_INFLOW, 0) AS AMOUNT_INFLOW, " +
                    "ISNULL(B.AMOUNT_OUTFLOW, 0) AS AMOUNT_OUTFLOW, ISNULL(A.AMOUNT_INFLOW, 0) -ISNULL(B.AMOUNT_OUTFLOW, 0) AS BALANCE FROM PLAN_ACCOUNT pa LEFT JOIN (SELECT CONVERT(char(10), pla.PAYMENT_DATE, 111) AS DATE_INFLOW, " +
                    "SUM(pla.AMOUNT) AS AMOUNT_INFLOW FROM PLAN_ACCOUNT pla WHERE ISDEBIT = 'Y' AND STATUS <> 0 AND PAYMENT_DATE BETWEEN CONVERT(char(10), getdate(), 111) AND getdate() + 180 " +
                    "GROUP BY CONVERT(char(10), PAYMENT_DATE, 111))A ON CONVERT(char(10), pa.PAYMENT_DATE, 111) = A.DATE_INFLOW LEFT JOIN (SELECT CONVERT(char(10), pla.PAYMENT_DATE, 111) AS DATE_OUTFLOW, " +
                    "SUM(pla.AMOUNT) AS AMOUNT_OUTFLOW FROM PLAN_ACCOUNT pla WHERE ISDEBIT = 'N' AND STATUS <> 0 AND PAYMENT_DATE BETWEEN CONVERT(char(10), getdate(), 111) AND getdate() + 180 GROUP BY CONVERT(char(10), PAYMENT_DATE, 111))B " +
                    "ON CONVERT(char(10), pa.PAYMENT_DATE, 111) = B.DATE_OUTFLOW)CASHFLOW_2 WHERE CASHFLOW_1.DATE_CASHFLOW >= CASHFLOW_2.DATE_CASHFLOW GROUP BY CASHFLOW_1.DATE_CASHFLOW,CASHFLOW_1.AMOUNT_INFLOW, CASHFLOW_1.AMOUNT_OUTFLOW,CASHFLOW_1.BALANCE)C " +
                    "WHERE C.AMOUNT_INFLOW <> 0 OR C.AMOUNT_OUTFLOW <> 0 ORDER BY C.DATE_CASHFLOW ASC ; ").ToList();
                logger.Info("Get Cash Flow Count=" + lstItem.Count);
            }

            return lstItem;
        }

        //取得特定日期之收入(Debit)
        public List<PLAN_ACCOUNT> getDebitByDate(string date)
        {
            List<PLAN_ACCOUNT> lstDebit = new List<PLAN_ACCOUNT>();
            using (var context = new topmepEntities())
            {
                lstDebit = context.PLAN_ACCOUNT.SqlQuery("select a.* from PLAN_ACCOUNT a "
                    + "where a.ISDEBIT = 'Y' AND CONVERT(char(10),a.PAYMENT_DATE,111) = @date "
                   , new SqlParameter("date", date)).ToList();
            }
            return lstDebit;
        }

        //取得特定日期之支出(Credit)
        public List<PLAN_ACCOUNT> getCreditByDate(string date)
        {
            List<PLAN_ACCOUNT> lstCredit = new List<PLAN_ACCOUNT>();
            using (var context = new topmepEntities())
            {
                lstCredit = context.PLAN_ACCOUNT.SqlQuery("select a.* from PLAN_ACCOUNT a "
                    + "where a.ISDEBIT = 'N' AND CONVERT(char(10),a.PAYMENT_DATE,111) = @date "
                   , new SqlParameter("date", date)).ToList();
            }
            return lstCredit;
        }

        #region 公司費用預算
        //公司費用預算上傳 
        public int refreshExpBudget(List<FIN_EXPENSE_BUDGET> items)
        {
            int i = 0;
            logger.Info("refreshExpBudgetItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (FIN_EXPENSE_BUDGET item in items)
                {
                    context.FIN_EXPENSE_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add FIN_EXPENSE_BUDGET count =" + i);
            return i;
        }
        public int delExpBudgetByYear(int year)
        {
            logger.Info("remove all expense budget by budget year=" + year);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all FIN_EXPENSE_BUDGET by budget year =" + year);
                i = context.Database.ExecuteSqlCommand("DELETE FROM FIN_EXPENSE_BUDGET WHERE BUDGET_YEAR=@year", new SqlParameter("@year", year));
            }
            logger.Debug("deleteFIN_EXPENSE_BUDGET count=" + i);
            return i;
        }

        //取得特定年度公司費用預算
        public List<ExpenseBudgetSummary> getExpBudgetByYear(int year)
        {
            List<ExpenseBudgetSummary> lstExpBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstExpBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM (SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM (SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM FIN_EXPENSE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE BUDGET_YEAR = @year) As STable " +
                   "PIVOT (SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC ORDER BY A.SUBJECT_ID ; "
                   , new SqlParameter("year", year)).ToList();
            }
            return lstExpBudget;
        }

        //取得公司費用科目代碼與名稱
        public List<FIN_SUBJECT> getExpBudgetSubject()
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE CATEGORY = '公司營業費用' ORDER BY FIN_SUBJECT_ID ; ").ToList();
            }
            return lstSubject;
        }

        //取得特定年度之公司費用總預算金額
        public ExpenseBudgetSummary getTotalExpBudgetAmount(int year)
        {
            ExpenseBudgetSummary lstAmount = null;
            using (var context = new topmepEntities())
            {
                lstAmount = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT  SUM(AMOUNT) AS TOTAL_BUDGET FROM FIN_EXPENSE_BUDGET WHERE BUDGET_YEAR = @year  "
               , new SqlParameter("year", year)).FirstOrDefault();
            }
            return lstAmount;
        }

        //取得特定年度之公司費用總執行金額
        public ExpenseBudgetSummary getTotalOperationExpAmount(int year)
        {
            ExpenseBudgetSummary lstExpAmount = null;
            using (var context = new topmepEntities())
            {
                lstExpAmount = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT SUM(ei.AMOUNT) AS TOTAL_OPERATION_EXP FROM FIN_EXPENSE_ITEM ei LEFT JOIN FIN_EXPENSE_FORM ef " +
                    "ON ei.EXP_FORM_ID = ef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fs.CATEGORY = '公司營業費用' " +
                    "AND ef.OCCURRED_YEAR = @year AND ef.OCCURRED_MONTH > 6 OR fs.CATEGORY = '公司營業費用' AND ef.OCCURRED_YEAR = @year + 1 AND ef.OCCURRED_MONTH < 7  "
               , new SqlParameter("year", year)).FirstOrDefault();
            }
            return lstExpAmount;
        }

        //取得特定年度公司費用預算與費用彙整
        public List<ExpenseBudgetSummary> getExpBudgetSummaryByYear(int year)
        {
            List<ExpenseBudgetSummary> lstExpBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstExpBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT ROW_NUMBER() OVER(ORDER BY D.SUBJECT_ID, D.SUB_NO) AS NO, D.* FROM (SELECT ROW_NUMBER() OVER(ORDER BY A.SUBJECT_ID) AS SUB_NO, A.*, " +
                    "SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                    "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                    "FROM(SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                    "FROM(SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM FIN_EXPENSE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE BUDGET_YEAR = @year) As STable " +
                    "PIVOT(SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                    "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC UNION " +
                    "SELECT ROW_NUMBER() OVER(ORDER BY C.FIN_SUBJECT_ID) + 1 AS SUB_NO, C.*, SUM(ISNULL(C.JAN, 0)) + SUM(ISNULL(C.FEB, 0)) + SUM(ISNULL(C.MAR, 0)) + SUM(ISNULL(C.APR, 0)) + SUM(ISNULL(C.MAY, 0)) + SUM(ISNULL(C.JUN, 0)) " +
                    "+ SUM(ISNULL(C.JUL, 0)) + SUM(ISNULL(C.AUG, 0)) + SUM(ISNULL(C.SEP, 0)) + SUM(ISNULL(C.OCT, 0)) + SUM(ISNULL(C.NOV, 0)) + SUM(ISNULL(C.DEC, 0)) AS HTOTAL " +
                    "FROM(SELECT SUBJECT_NAME, FIN_SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                    "FROM(SELECT B.OCCURRED_MONTH, fs.FIN_SUBJECT_ID, fs.SUBJECT_NAME, B.AMOUNT FROM FIN_SUBJECT fs LEFT JOIN(SELECT ef.OCCURRED_MONTH, ei.FIN_SUBJECT_ID, ei.AMOUNT FROM FIN_EXPENSE_ITEM ei " +
                    "LEFT JOIN FIN_EXPENSE_FORM ef ON ei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.OCCURRED_YEAR = @year AND ef.OCCURRED_MONTH > 6 OR ef.OCCURRED_YEAR = @year + 1 AND ef.OCCURRED_MONTH < 7)B " +
                    "ON fs.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID WHERE fs.CATEGORY = '公司營業費用') As STable " +
                    "PIVOT(SUM(AMOUNT) FOR OCCURRED_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)C " +
                    "GROUP BY C.SUBJECT_NAME, C.FIN_SUBJECT_ID, C.JAN, C.FEB, C.MAR, C.APR, C.MAY, C.JUN, C.JUL, C.AUG, C.SEP, C.OCT, C.NOV, C.DEC )D ORDER BY D.SUBJECT_ID, D.SUB_NO ; "
                   , new SqlParameter("year", year)).ToList();
            }
            return lstExpBudget;
        }
        #endregion

        #region 公司費用
        public ExpenseFormFunction formEXP = null;
        public List<ExpenseBudgetSummary> EXPItem = null;
        public List<ExpenseBudgetSummary> siteEXPItem = null;
        //取得公司費用項目
        public List<FIN_SUBJECT> getSubjectOfExpense()
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE CATEGORY = '公司營業費用' ORDER BY FIN_SUBJECT_ID ; ").ToList();
                logger.Info("Get Subject of Operating Expense Count=" + lstSubject.Count);
            }
            return lstSubject;
        }

        //取得特定費用項目
        public List<FIN_SUBJECT> getSubjectByChkItem(string[] lstItemId)
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
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
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE FIN_SUBJECT_ID IN (" + ItemId + ") ; ").ToList();
                logger.Info("Get Subject of Expense  Count=" + lstSubject.Count);
            }
            return lstSubject;
        }

        public string newExpenseForm(FIN_EXPENSE_FORM form)
        {
            //1.建立公司營業費用單/工地費用單
            logger.Info("create new expense form ");
            string sno_key = "EXP";
            SerialKeyService snoservice = new SerialKeyService();
            form.EXP_FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new expense form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.FIN_EXPENSE_FORM.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add form=" + i);
                logger.Info("expense form id = " + form.EXP_FORM_ID);
                //if (i > 0) { status = true; };
            }
            return form.EXP_FORM_ID;
        }

        public int AddExpenseItems(List<FIN_EXPENSE_ITEM> lstItem)
        {
            //2.新增費用項目資料
            int j = 0;
            logger.Info("add expense items = " + lstItem.Count);
            using (var context = new topmepEntities())
            {
                //3.將費用項目資料寫入 
                foreach (FIN_EXPENSE_ITEM item in lstItem)
                {
                    context.FIN_EXPENSE_ITEM.Add(item);
                }

                j = context.SaveChanges();
                logger.Info("add expense count =" + j);
            }
            return j;
        }

        //取得公司費用單/工地費用單
        public void getEXPByExpId(string expid)
        {
            logger.Info("get form : formid=" + expid);
            using (var context = new topmepEntities())
            {
                //取得費用單檔頭資訊
                string sql = "SELECT fef.EXP_FORM_ID, fef.PROJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.PAYMENT_DATE, fef.STATUS, fef.CREATE_ID, fef.CREATE_DATE, fef.REMARK, fef.MODIFY_DATE, fef.PASS_CREATE_ID, fef.PASS_CREATE_DATE, fef.APPROVE_CREATE_ID, fef.APPROVE_CREATE_DATE, " +
                    "fef.JOURNAL_CREATE_ID, fef.JOURNAL_CREATE_DATE, p.PROJECT_NAME FROM FIN_EXPENSE_FORM fef LEFT JOIN TND_PROJECT p ON fef.PROJECT_ID = p.PROJECT_ID WHERE fef.EXP_FORM_ID =@expid ";
                formEXP = context.Database.SqlQuery<ExpenseFormFunction>(sql, new SqlParameter("expid", expid)).First();
                //取得公司營業費用單明細
                EXPItem = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT G.*, C.CUM_BUDGET, D.CUM_YEAR_AMOUNT, G.AMOUNT / ISNULL(G.BUDGET_AMOUNT, 1) *100 AS MONTH_RATIO, D.CUM_YEAR_AMOUNT / ISNULL(C.CUM_BUDGET,1) *100 AS YEAR_RATIO " +
                    "FROM (SELECT A.*, B.AMOUNT AS BUDGET_AMOUNT FROM (SELECT fei.*, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.STATUS, fs.SUBJECT_NAME FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs " +
                    "ON fei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fei.EXP_FORM_ID = @expid)A " +
                    "LEFT JOIN (SELECT F.*, feb.AMOUNT FROM (SELECT fei.FIN_SUBJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs  " +
                    "ON fei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fei.EXP_FORM_ID = @expid)F LEFT JOIN FIN_EXPENSE_BUDGET feb ON F.FIN_SUBJECT_ID + CONVERT(varchar, OCCURRED_YEAR) + CONVERT(varchar, OCCURRED_MONTH) = feb.SUBJECT_ID + CONVERT(varchar, CURRENT_YEAR) + CONVERT(varchar, BUDGET_MONTH))B ON A.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID)G " +
                    "LEFT JOIN (SELECT SUBJECT_ID, SUM(AMOUNT) AS CUM_BUDGET FROM FIN_EXPENSE_BUDGET WHERE CURRENT_YEAR = " + formEXP.OCCURRED_YEAR + "AND BUDGET_MONTH <= " + formEXP.OCCURRED_MONTH + "GROUP BY SUBJECT_ID) C ON G.FIN_SUBJECT_ID = C.SUBJECT_ID " +
                    "LEFT JOIN (SELECT fei.FIN_SUBJECT_ID, SUM(AMOUNT) AS CUM_YEAR_AMOUNT FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID " +
                    "WHERE fef.OCCURRED_YEAR = " + formEXP.OCCURRED_YEAR + "AND fef.OCCURRED_MONTH <= " + formEXP.OCCURRED_MONTH + "GROUP BY fei.FIN_SUBJECT_ID)D ON G.FIN_SUBJECT_ID = D.FIN_SUBJECT_ID ", new SqlParameter("expid", expid)).ToList();
                logger.Debug("get query year of operating expense:" + formEXP.OCCURRED_YEAR);
                logger.Debug("get operating expense item count:" + EXPItem.Count);
                //取得工地費用單明細
                siteEXPItem = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT G.*, C.CUM_BUDGET, D.CUM_YEAR_AMOUNT, G.AMOUNT / G.BUDGET_AMOUNT *100 AS MONTH_RATIO, D.CUM_YEAR_AMOUNT / C.CUM_BUDGET *100 AS YEAR_RATIO " +
                    "FROM (SELECT A.*, B.AMOUNT AS BUDGET_AMOUNT FROM (SELECT fei.*, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.STATUS, fs.SUBJECT_NAME FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs " +
                    "ON fei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fei.EXP_FORM_ID = @expid)A " +
                    "LEFT JOIN (SELECT F.*, sb.AMOUNT FROM (SELECT fei.FIN_SUBJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs  " +
                    "ON fei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fei.EXP_FORM_ID = @expid)F LEFT JOIN PLAN_SITE_BUDGET sb ON F.FIN_SUBJECT_ID + CONVERT(varchar, OCCURRED_YEAR) + CONVERT(varchar, OCCURRED_MONTH) = sb.SUBJECT_ID + CONVERT(varchar, BUDGET_YEAR) + CONVERT(varchar, BUDGET_MONTH))B ON A.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID)G " +
                    "LEFT JOIN (SELECT SUBJECT_ID, SUM(AMOUNT) AS CUM_BUDGET FROM PLAN_SITE_BUDGET WHERE BUDGET_YEAR = " + formEXP.OCCURRED_YEAR + "AND BUDGET_MONTH <= " + formEXP.OCCURRED_MONTH + "GROUP BY SUBJECT_ID) C ON G.FIN_SUBJECT_ID = C.SUBJECT_ID " +
                    "LEFT JOIN (SELECT fei.FIN_SUBJECT_ID, SUM(AMOUNT) AS CUM_YEAR_AMOUNT FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID " +
                    "WHERE fef.OCCURRED_YEAR = " + formEXP.OCCURRED_YEAR + "AND fef.OCCURRED_MONTH <= " + formEXP.OCCURRED_MONTH + "GROUP BY fei.FIN_SUBJECT_ID)D ON G.FIN_SUBJECT_ID = D.FIN_SUBJECT_ID ", new SqlParameter("expid", expid)).ToList();
                logger.Debug("get query year of plan site expense:" + formEXP.OCCURRED_YEAR);
                logger.Debug("get plan site expense item count:" + siteEXPItem.Count);
            }
        }

        public int refreshEXPForm(string formid, FIN_EXPENSE_FORM ef, List<FIN_EXPENSE_ITEM> lstItem)
        {
            logger.Info("Update expense form id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(ef).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update expense form =" + i);
                    logger.Info("expense form item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (FIN_EXPENSE_ITEM item in lstItem)
                    {
                        FIN_EXPENSE_ITEM existItem = null;
                        logger.Debug("form item id=" + item.EXP_ITEM_ID);
                        if (item.EXP_ITEM_ID != 0)
                        {
                            existItem = context.FIN_EXPENSE_ITEM.Find(item.EXP_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.FIN_SUBJECT_ID));
                            string sql = "SELECT * FROM FIN_EXPENSE_ITEM WHERE EXP_FORM_ID=@formid AND FIN_SUBJECT_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",fin_subject_id=" + item.FIN_SUBJECT_ID);
                            FIN_EXPENSE_ITEM excelItem = context.FIN_EXPENSE_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.FIN_EXPENSE_ITEM.Find(excelItem.EXP_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.FIN_SUBJECT_ID);
                        existItem.AMOUNT = item.AMOUNT;
                        existItem.ITEM_REMARK = item.ITEM_REMARK;
                        context.FIN_EXPENSE_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update expense form item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new expense form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //更新費用單狀態為送審
        public int RefreshEXPStatusById(string expid)
        {
            int i = 0;
            logger.Info("update the status of EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 20 WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //取得符合條件之公司營業費用單名單
        public List<OperatingExpenseFunction> getEXPListByExpId(string occurreddate, string subjectname, string expid, int status, string projectid)
        {
            logger.Info("search expense form by " + occurreddate + ", 費用單編號 =" + expid + ", 項目名稱 =" + subjectname + ", 估驗單狀態 =" + status + ", 專案編號 =" + projectid);
            List<OperatingExpenseFunction> lstForm = new List<OperatingExpenseFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            if (15 == status)//預設狀態(STATUS >10 AND <20, 即狀態為退件或草稿)
            {
                string sql = "SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*, (SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B ";

                sql = sql + "WHERE B.STATUS > 10 AND ISNULL(B.PROJECT_ID, '') =@id ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", projectid));
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else if (20 == status)
            {
                string sql = "SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*,(SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B ";

                sql = sql + "WHERE B.STATUS = 20 AND ISNULL(B.PROJECT_ID, '') =@id ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", projectid));
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else if (30 == status)
            {
                string sql = "SELECT ISNULL(B.PROJECT_ID,'公司營業費用') AS PROJECT_ID, B.PROJECT_NAME, B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*, p.PROJECT_NAME, (SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef LEFT JOIN TND_PROJECT p ON ef.PROJECT_ID = p.PROJECT_ID)B ";

                sql = sql + "WHERE B.STATUS = 30 ";
                var parameters = new List<SqlParameter>();
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else if (40 == status)
            {
                string sql = "SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*,(SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B ";

                sql = sql + "WHERE B.STATUS = 40 AND ISNULL(B.PROJECT_ID, '') =@id ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", projectid));
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else
            {
                using (var context = new topmepEntities())
                {
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>("SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*,(SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B WHERE B.STATUS < 20 AND ISNULL(B.PROJECT_ID, '') =@id ", new SqlParameter("id", projectid)).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            return lstForm;
        }

        //更新公司營業費用單狀態為退件
        public int RejectEXPByExpId(string expid)
        {
            int i = 0;
            logger.Info("reject EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 0 WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新公司營業費用單狀態為主管已通過
        public int PassEXPByExpId(string expid, string passid)
        {
            int i = 0;
            logger.Info("Pass EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 30, PASS_CREATE_ID = @passid, PASS_CREATE_DATE = GETDATE() WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            parameters.Add(new SqlParameter("passid", passid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新公司營業費用單狀態為已立帳(即會計已完成稽核)
        public int JournalByExpId(string expid, string journalid)
        {
            int i = 0;
            logger.Info("Journal For EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 40, JOURNAL_CREATE_ID = @journalid, JOURNAL_CREATE_DATE = GETDATE() WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            parameters.Add(new SqlParameter("journalid", journalid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新公司營業費用單狀態為已核可
        public int ApproveEXPByExpId(string expid, string approveid)
        {
            int i = 0;
            logger.Info("Approve EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 50, APPROVE_CREATE_ID = @approveid, APPROVE_CREATE_DATE = GETDATE() WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            parameters.Add(new SqlParameter("approveid", approveid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }
        public string AddAccountByExpId(string formid, string createid)
        {
            //寫入現金流支出資料
            using (var context = new topmepEntities())
            {
                List<PLAN_ACCOUNT> lstItem = new List<PLAN_ACCOUNT>();

                string sql = "INSERT INTO PLAN_ACCOUNT (ACCOUNT_FORM_ID, PAYMENT_DATE, AMOUNT, "
                    + "ACCOUNT_TYPE, ISDEBIT, STATUS, CREATE_ID) "
                    + "SELECT ef.EXP_FORM_ID AS ACCOUNT_FORM_ID, ef.PAYMENT_DATE, SUM(ei.AMOUNT) AS AMOUNT, 'O' AS ACCOUNT_TYPE, 'N' AS ISDEBIT, 10 AS STATUS, '" + createid + "' AS CREATE_ID "
                    + "FROM FIN_EXPENSE_FORM ef LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID " +
                    "WHERE ef.EXP_FORM_ID = '" + formid + "' GROUP BY ef.EXP_FORM_ID, ef.PAYMENT_DATE ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return formid;
            }
        }

        public int refreshAccountStatus(List<PLAN_ACCOUNT> lstItem)
        {
            int j = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("plan account item = " + lstItem.Count);
                //將item資料寫入 
                foreach (PLAN_ACCOUNT item in lstItem)
                {
                    PLAN_ACCOUNT existItem = null;
                    logger.Debug("form item id=" + item.PLAN_ACCOUNT_ID);
                    if (item.PLAN_ACCOUNT_ID != 0)
                    {
                        existItem = context.PLAN_ACCOUNT.Find(item.PLAN_ACCOUNT_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("itemid", item.ACCOUNT_FORM_ID));
                        string sql = "SELECT * FROM PLAN_ACCOUNT WHERE ACCOUNT_FORM_ID=@itemid";
                        logger.Info(sql + " ;" + item.ACCOUNT_FORM_ID);
                        PLAN_ACCOUNT excelItem = context.PLAN_ACCOUNT.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ACCOUNT.Find(excelItem.PLAN_ACCOUNT_ID);

                    }
                    logger.Debug("find exist item=" + existItem.ACCOUNT_FORM_ID);
                    existItem.STATUS = item.STATUS;
                    context.PLAN_ACCOUNT.AddOrUpdate(existItem);
                }
                j = context.SaveChanges();
                logger.Debug("Update plan account item =" + j);
            }
            return j;
        }

        //取得符合條件之帳款資料
        public List<PlanAccountFunction> getPlanAccount(string paymentdate, string projectname, string payee, string accounttype)
        {
            logger.Info("search plan account by " + paymentdate + ", 受款人 =" + payee + ", 專案名稱 =" + projectname + ", 帳款類型 =" + accounttype);
            List<PlanAccountFunction> lstForm = new List<PlanAccountFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pa.AMOUNT, pa.ACCOUNT_TYPE, CONVERT(char(10), pa.PAYMENT_DATE, 111) AS RECORDED_DATE, pa.PLAN_ACCOUNT_ID, pa.STATUS, p.PROJECT_NAME, s.COMPANY_NAME AS PAYEE FROM PLAN_ACCOUNT pa LEFT JOIN TND_PROJECT p ON pa.PROJECT_ID = p.PROJECT_ID " +
                "LEFT JOIN TND_SUPPLIER s ON SUBSTRING(pa.CONTRACT_ID, 7, 7) = s.SUPPLIER_ID WHERE pa.ACCOUNT_TYPE = @accounttype ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("accounttype", accounttype));
            logger.Info("sql=" + sql);

            //支付日期條件
            if (null != paymentdate && paymentdate != "")
            {
                sql = sql + "AND CONVERT(char(10), pa.PAYMENT_DATE, 111) =@paymentdate ";
                parameters.Add(new SqlParameter("paymentdate", paymentdate));
            }
            //專案名稱條件
            if (null != projectname && projectname != "")
            {
                sql = sql + "AND p.PROJECT_NAME LIKE @projectname ";
                parameters.Add(new SqlParameter("projectname", '%' + projectname + '%'));
            }
            //受款人條件
            if (null != payee && payee != "")
            {
                sql = sql + "AND s.COMPANY_NAME LIKE @payee ";
                parameters.Add(new SqlParameter("payee", '%' + payee + '%'));
            }
            sql = sql + "ORDER BY pa.PAYMENT_DATE DESC ";
            using (var context = new topmepEntities())
            {
                logger.Debug("get plan account sql=" + sql);
                lstForm = context.Database.SqlQuery<PlanAccountFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("get plan account count=" + lstForm.Count);
            }
            return lstForm;
        }

        public PlanAccountFunction getPlanAccountItem(string itemid)
        {
            logger.Debug("get plan account item by id=" + itemid);
            PlanAccountFunction aitem = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                aitem = context.Database.SqlQuery<PlanAccountFunction>("SELECT PARSENAME(Convert(varchar,Convert(money,pa.AMOUNT),1),2) AS RECORDED_AMOUNT, pa.ACCOUNT_TYPE, CONVERT(char(10), pa.PAYMENT_DATE, 111) AS RECORDED_DATE, " +
                    "pa.PLAN_ACCOUNT_ID, pa.CONTRACT_ID, pa.ACCOUNT_TYPE, pa.ACCOUNT_FORM_ID, pa.ISDEBIT, pa.STATUS, pa.CREATE_ID, p.PROJECT_NAME FROM PLAN_ACCOUNT pa LEFT JOIN TND_PROJECT p ON pa.PROJECT_ID = p.PROJECT_ID " +
                    "WHERE pa.PLAN_ACCOUNT_ID=@itemid ",
                new SqlParameter("itemid", itemid)).First();
            }
            return aitem;
        }

        public int updatePlanAccountItem(PLAN_ACCOUNT item)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_ACCOUNT.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("updatePlanAcountItem  fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }

        #endregion

        #region 工地費用

        //取得工地費用項目
        public List<FIN_SUBJECT> getSubjectOfExpense4Site()
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE CATEGORY = '工地費用' ORDER BY FIN_SUBJECT_ID; ").ToList();
                logger.Info("Get Subject of Operating Expense Count=" + lstSubject.Count);
            }
            return lstSubject;
        }

        public string getSiteBudgetById(string prjid)
        {
            string projectid = null;
            using (var context = new topmepEntities())
            {
                projectid = context.Database.SqlQuery<string>("SELECT DISTINCT PROJECT_ID FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return projectid;
        }

        public int delSiteBudgetByProject(string projectid, string year)
        {
            logger.Info("remove all site budget by projectid =" + projectid + "and by year sequence =" + year);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all PLAN_SITE_BUDGET by projectid =" + projectid + "and by year sequence =" + year);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_SITE_BUDGET WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE =@year", new SqlParameter("@projectid", projectid), new SqlParameter("@year", year));
            }
            logger.Debug("delete PALN_SITE_BUDGET count=" + i);
            return i;
        }

        public int refreshSiteBudget(List<PLAN_SITE_BUDGET> items)
        {
            int i = 0;
            logger.Info("refreshSiteBudgetItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_SITE_BUDGET item in items)
                {
                    context.PLAN_SITE_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add PLAN_SITE_BUDGET count =" + i);
            return i;
        }

        //取得專案工地費用預算
        #region 第1年度
        public List<ExpenseBudgetSummary> getFirstYearBudgetByProject(string projectid)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM (SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM (SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM PLAN_SITE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE = '1') As STable " +
                   "PIVOT (SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC ORDER BY A.SUBJECT_ID ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lstSiteBudget;
        }
        #endregion
        #region 第2年度
        public List<ExpenseBudgetSummary> getSecondYearBudgetByProject(string projectid)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM (SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM (SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM PLAN_SITE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE = '2') As STable " +
                   "PIVOT (SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC ORDER BY A.SUBJECT_ID ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lstSiteBudget;
        }
        #endregion
        #region 第3年度
        public List<ExpenseBudgetSummary> getThirdYearBudgetByProject(string projectid)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM (SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM (SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM PLAN_SITE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE = '3') As STable " +
                   "PIVOT (SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC ORDER BY A.SUBJECT_ID ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lstSiteBudget;
        }
        #endregion
        #region 第4年度
        public List<ExpenseBudgetSummary> getFourthYearBudgetByProject(string projectid)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM (SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM (SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM PLAN_SITE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE = '4') As STable " +
                   "PIVOT (SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC ORDER BY A.SUBJECT_ID ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lstSiteBudget;
        }
        #endregion
        #region 第5年度
        public List<ExpenseBudgetSummary> getFifthYearBudgetByProject(string projectid)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM (SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM (SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM PLAN_SITE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE = '5') As STable " +
                   "PIVOT (SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC ORDER BY A.SUBJECT_ID ; "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lstSiteBudget;
        }
        #endregion

        //取得專案工地費用預算之西元年
        #region 第1年度
        public int getFirstYearOfSiteBudgetById(string prjid)
        {
            int year = 0;
            using (var context = new topmepEntities())
            {
                year = context.Database.SqlQuery<int>("SELECT DISTINCT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid AND YEAR_SEQUENCE = '1' "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return year;
        }
        #endregion
        #region 第2年度
        public int getSecondYearOfSiteBudgetById(string prjid)
        {
            int year = 0;
            using (var context = new topmepEntities())
            {
                year = context.Database.SqlQuery<int>("SELECT DISTINCT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid AND YEAR_SEQUENCE = '2' "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return year;
        }
        #endregion
        #region 第3年度
        public int getThirdYearOfSiteBudgetById(string prjid)
        {
            int year = 0;
            using (var context = new topmepEntities())
            {
                year = context.Database.SqlQuery<int>("SELECT DISTINCT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid AND YEAR_SEQUENCE = '3' "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return year;
        }
        #endregion
        #region 第4年度
        public int getFourthYearOfSiteBudgetById(string prjid)
        {
            int year = 0;
            using (var context = new topmepEntities())
            {
                year = context.Database.SqlQuery<int>("SELECT DISTINCT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid AND YEAR_SEQUENCE = '4' "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return year;
        }
        #endregion
        #region 第5年度
        public int getFifthYearOfSiteBudgetById(string prjid)
        {
            int year = 0;
            using (var context = new topmepEntities())
            {
                year = context.Database.SqlQuery<int>("SELECT DISTINCT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid AND YEAR_SEQUENCE = '5' "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return year;
        }
        #endregion
        //取得特定專案之工地費用總預算金額
        public ExpenseBudgetSummary getTotalSiteBudgetAmount(string projectid)
        {
            ExpenseBudgetSummary lstAmount = null;
            using (var context = new topmepEntities())
            {
                lstAmount = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT  SUM(AMOUNT) AS TOTAL_BUDGET FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @projectid  "
               , new SqlParameter("projectid", projectid)).FirstOrDefault();
            }
            return lstAmount;
        }
        #endregion
    }
}
