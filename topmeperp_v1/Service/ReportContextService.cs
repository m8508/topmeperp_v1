using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using topmeperp.Models;
using Newtonsoft.Json.Linq;

namespace topmeperp.Service
{
    #region 比價資料處理
    public class RptCompareProjectPrice : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public RptCompareProjectPrice()
        {
        }
        //依據新標單資料，取得相關專案的舊報價資料
        public List<ProjectCompareData> RtpGetPriceFromExistProject(string srcProjectId, string tarProjectId, bool hasProject, bool hasPrice)
        {
            List<ProjectCompareData> lstCompareData = null;
            try
            {
                using (var context = new topmepEntities())
                {
                    string sql = "SELECT DISTINCT SRC.PROJECT_ID SOURCE_PROJECT_ID,"
                        + "SRC.SYSTEM_MAIN SOURCE_SYSTEM_MAIN,"
                        + "SRC.SYSTEM_SUB SOURCE_SYSTEM_SUB,"
                        + "SRC.ITEM_ID SOURCE_ITEM_ID,"
                        + "SRC.ITEM_DESC SOURCE_ITEM_DESC,"
                        + "SRC.ITEM_UNIT_PRICE SRC_UNIT_PRICE,"
                        + "TAR.ITEM_ID TARGET_ITEM_ID,"
                        + "TAR.ITEM_DESC TARGET_ITEM_DESC,"
                        + "TAR.SYSTEM_MAIN TARGET_SYSTEM_MAIN,"
                        + "TAR.SYSTEM_SUB TARGET_SYSTEM_SUB,"
                        + "TAR.PROJECT_ID TARGET_PROJECT_ID,"
                        + "SRC.EXCEL_ROW_ID EXCEL_ROW_ID FROM "
                        + "(SELECT PROJECT_ID, ITEM_ID, ITEM_DESC, ITEM_UNIT_PRICE, SYSTEM_MAIN, ISNULL(SYSTEM_SUB,'*') SYSTEM_SUB, EXCEL_ROW_ID FROM TND_PROJECT_ITEM WHERE PROJECT_ID=@srcProjectId ) SRC,"
                        + "(SELECT PROJECT_ID, ITEM_ID, ITEM_DESC, SYSTEM_MAIN, ISNULL(SYSTEM_SUB,'*') SYSTEM_SUB FROM TND_PROJECT_ITEM WHERE PROJECT_ID=@tarProjectId ) TAR "
                        + "WHERE SRC.ITEM_DESC = TAR.ITEM_DESC ";

                    if (hasProject)
                    {
                        sql = sql + "AND SRC.SYSTEM_MAIN = TAR.SYSTEM_MAIN AND SRC.SYSTEM_SUB = TAR.SYSTEM_SUB  ";
                    }
                    if (hasPrice)
                    {
                        sql = sql + "AND ITEM_UNIT_PRICE is not null ";
                    }
                    sql = sql + "ORDER BY EXCEL_ROW_ID;";

                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("srcProjectId", srcProjectId));
                    parameters.Add(new SqlParameter("tarProjectId", tarProjectId));
                    logger.Info("SQL=" + sql);
                    lstCompareData = context.Database.SqlQuery<ProjectCompareData>(sql, parameters.ToArray()).ToList();
                    logger.Info("Get CompareData Record Count=" + lstCompareData.Count);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
            }
            return lstCompareData;
        }
        public int MigratePrice(List<ProjectCompareData> lstData)
        {
            int i = 0;
            try
            {
                using (var context = new topmepEntities())
                {
                    string sql = "UPDATE TND_PROJECT_ITEM SET ITEM_UNIT_PRICE=@price WHERE PROJECT_ID=@projectid AND SYSTEM_MAIN=@systemMain AND ISNULL(SYSTEM_SUB,'*')=@systemSub AND ITEM_DESC=@itemDesc;";
                    logger.Info("sql=" + sql);
                    foreach (ProjectCompareData data in lstData)
                    {
                        //item.SOURCE_PROJECT_ID + '|' + item.SOURCE_SYSTEM_MAIN + '|' + item.SOURCE_SYSTEM_SUB + '|' + item.SRC_UNIT_PRICE + '|' + item.TARGET_PROJECT_ID + '|' + item.SOURCE_ITEM_DESC;}
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("price", data.SRC_UNIT_PRICE));
                        logger.Debug("price=" + data.SRC_UNIT_PRICE);
                        parameters.Add(new SqlParameter("projectid", data.TARGET_PROJECT_ID));
                        logger.Debug("TARGET_PROJECT_ID=" + data.TARGET_PROJECT_ID);
                        parameters.Add(new SqlParameter("systemMain", data.SOURCE_SYSTEM_MAIN));
                        logger.Debug("SOURCE_SYSTEM_MAIN=" + data.SOURCE_SYSTEM_MAIN);
                        parameters.Add(new SqlParameter("systemSub", data.SOURCE_SYSTEM_SUB));
                        logger.Debug("SOURCE_SYSTEM_SUB=" + data.SOURCE_SYSTEM_SUB);
                        parameters.Add(new SqlParameter("itemDesc", data.TARGET_ITEM_DESC));
                        logger.Debug("TARGET_ITEM_DESC=" + data.TARGET_ITEM_DESC);
                        i = i + context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                        context.SaveChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
            }
            return i;
        }
    }
    #endregion
    #region 專案進度管理
    public class ProjectPlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        /// <summary>
        /// 取得專案所有任務
        /// </summary>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public DataTable getProjectTask(string projectid)
        {
            string sql = "WITH PrjTree(TASK_NAME, PRJ_UID, LV_NO,PRJ_ID,PARENT_UID,START_DATE,FINISH_DATE,DURATION) AS "
                + " (SELECT TASK_NAME, PRJ_UID, 0 LV_NO,PRJ_ID ,Null,START_DATE,FINISH_DATE,DURATION "
                + " FROM PLAN_TASK  WHERE PARENT_UID IS  NULL  AND PROJECT_ID = @projectid AND PRJ_UID = 0 "
                + " UNION ALL "
                + " SELECT P.TASK_NAME, P.PRJ_UID, B.LV_NO + 1,P.PRJ_ID,P.PARENT_UID,P.START_DATE,P.FINISH_DATE,P.DURATION "
                + " FROM PLAN_TASK P, PrjTree B "
                + " WHERE  P.PROJECT_ID=@projectid AND P.PARENT_UID = B.PRJ_UID and P.TASK_NAME is not null )"
                + " SELECT(REPLICATE('**', LV_NO) + TASK_NAME) as 'TASK_NAME',LV_NO,PRJ_UID,PARENT_UID,START_DATE,FINISH_DATE,DURATION "
                + " FROM PrjTree ORDER BY PRJ_ID";
            var parameters = new Dictionary<string, Object>();
            //設定專案名編號資料
            parameters.Add("projectid", projectid);
            logger.Debug("sql=" + sql);
            logger.Debug("prj_id=" + projectid);
            DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
            return ds.Tables[0];
        }
        /// <summary>
        /// 取得特定任務底下所有任務
        /// </summary>
        /// <param name="projectid"></param>
        /// <param name="prjuid"></param>
        /// <returns></returns>
        public DataTable getChildTask(string projectid, int prjuid)
        {
            string sql = "WITH PrjTree(TASK_NAME, PRJ_UID, LV_NO,PRJ_ID,PARENT_UID,START_DATE,FINISH_DATE,DURATION) AS "
                + " (SELECT TASK_NAME, PRJ_UID, 0 LV_NO,PRJ_ID,PARENT_UID,START_DATE,FINISH_DATE,DURATION "
                + " FROM PLAN_TASK  WHERE PROJECT_ID = @projectid AND PRJ_UID = @prjuid "
                + " UNION ALL "
                + " SELECT P.TASK_NAME, P.PRJ_UID, B.LV_NO + 1,P.PRJ_ID,P.PARENT_UID,P.START_DATE,P.FINISH_DATE,P.DURATION "
                + " FROM PLAN_TASK P, PrjTree B "
                + " WHERE P.PROJECT_ID=@projectid AND  P.PARENT_UID = B.PRJ_UID and P.TASK_NAME is not null )"
                + " SELECT(REPLICATE('**', LV_NO) + TASK_NAME) as 'TASK_NAME',LV_NO,PRJ_UID,PARENT_UID,START_DATE,FINISH_DATE,DURATION "
                + " FROM PrjTree ORDER BY PRJ_ID";
            logger.Debug("sql=" + sql);
            logger.Debug("prj_id=" + projectid + ",prjUID=" + prjuid);

            var parameters = new Dictionary<string, Object>();
            //設定專案名編號資料
            parameters.Add("projectid", projectid);
            parameters.Add("prjuid", prjuid);
            DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
            return ds.Tables[0];
        }
        public PLAN_TASK getRootTask(string projectid)
        {
            PLAN_TASK task = null;
            using (var context = new topmepEntities())
            {
                string sql = "SELECT * FROM PLAN_TASK WHERE PROJECT_ID=@projectid AND ROOT_TAG='Y';";
                task = context.PLAN_TASK.SqlQuery(sql, new SqlParameter("projectid", projectid)).First();
            }

            return task;
        }
        /// <summary>
        /// 設定專案任務起始任務
        /// </summary>
        /// <param name="projectid"></param>
        /// <param name="prjuid"></param>
        /// <returns></returns>
        public int setRootTask(string projectid, string prjuid)
        {
            int i = -1;
            using (var context = new topmepEntities())
            {
                string sql = "UPDATE PLAN_TASK SET ROOT_TAG=null WHERE PROJECT_ID=@projectid;";
                sql = sql + "UPDATE PLAN_TASK SET ROOT_TAG='Y' WHERE PROJECT_ID=@projectid AND PRJ_UID=@prjuid;";
                logger.Debug("sql=" + sql + ",projectid=" + projectid + ",prjuid=" + prjuid);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                logger.Debug("update row count=" + i);
            }
            return i;
        }
        /// <summary>
        /// 取得資料for tree
        /// </summary>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public void getProjectTask4Tree(string projectid)
        {
            string sql = "SELECT * FROM PLAN_TASK WHERE PROJECT_ID = @projectid and PRJ_ID>= "
                + "(SELECT PRJ_ID FROM PLAN_TASK where PROJECT_ID = @projectid and ROOT_TAG = 'Y') ORDER BY PRJ_ID;";
            List<PLAN_TASK> lstTask = new List<PLAN_TASK>();
            using (var context = new topmepEntities())
            {
                lstTask = context.PLAN_TASK.SqlQuery(sql, new SqlParameter("projectid", projectid)).ToList();
                logger.Debug("row count=" + lstTask.Count);
            }
            LinkedList<LinkedTree<PLAN_TASK>> lstTree = new LinkedList<LinkedTree<PLAN_TASK>>();
            foreach (PLAN_TASK t in lstTask)
            {
                LinkedTree<PLAN_TASK> node = new LinkedTree<PLAN_TASK>(t);

            }
        }
        public string convertToJson(DataTable dt)
        {
            JObject jo = new JObject();
            foreach (DataRow dr in dt.Rows)
            {

            }
            return null;
        }
    }
    #endregion

}
