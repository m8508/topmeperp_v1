using log4net;
using Newtonsoft.Json;
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

    #region 專案進度管理
    public class ProjectPlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //專案任務調用圖算數量物件
        MapInfoModels viewModel = new MapInfoModels();
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
                try
                {
                    string sql = "SELECT * FROM PLAN_TASK WHERE PROJECT_ID=@projectid AND ROOT_TAG='Y';";
                    logger.Debug("sql=" + sql + ",projectid=" + projectid);
                    task = context.PLAN_TASK.SqlQuery(sql, new SqlParameter("projectid", projectid)).First();
                }
                catch (Exception ex)
                {
                    logger.Error("Task Not found!!" + ex.Message);
                    logger.Error(ex.StackTrace);
                }
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
        public string getProjectTask4Tree(string projectid)
        {
            string sql = "SELECT * FROM PLAN_TASK WHERE PROJECT_ID = @projectid and PRJ_ID>= "
                + "(SELECT PRJ_ID FROM PLAN_TASK where PROJECT_ID = @projectid and ROOT_TAG = 'Y') ORDER BY PRJ_ID;";
            List<PLAN_TASK> lstTask = new List<PLAN_TASK>();
            using (var context = new topmepEntities())
            {
                lstTask = context.PLAN_TASK.SqlQuery(sql, new SqlParameter("projectid", projectid)).ToList();
                logger.Debug("row count=" + lstTask.Count);
            }
            // Dictionary<int, PROJECT_TASK_TREE_NODE> dicTree = new Dictionary<int, PROJECT_TASK_TREE_NODE>();
            //PROJECT_TASK_TREE_NODE rootnode = new PROJECT_TASK_TREE_NODE();
            Dictionary<int, TASK_TREE4SHOW> dicTree = new Dictionary<int, TASK_TREE4SHOW>();
            TASK_TREE4SHOW rootnode = new TASK_TREE4SHOW();
            foreach (PLAN_TASK t in lstTask)
            {
                //將跟節點置入Directory 內
                if (t.PARENT_UID == 0)
                {
                    //rootnode.tags.Add("工期:" + t.DURATION);
                    rootnode.tags.Add("完成:" + t.FINISH_DATE.Value.ToString("yyyy/MM/dd"));
                    rootnode.tags.Add("開始:" + t.START_DATE.Value.ToString("yyyy/MM/dd"));
                    rootnode.href = "#" + t.PRJ_UID;
                    rootnode.text = t.TASK_NAME;
                    dicTree.Add(t.PRJ_UID.Value, rootnode);
                    logger.Info("add root node :" + t.PRJ_UID);
                }
                else
                {
                    //將Dic 內的節點翻出，加入子節點
                    TASK_TREE4SHOW parentnode = (TASK_TREE4SHOW)dicTree[t.PARENT_UID.Value];
                    TASK_TREE4SHOW node = new TASK_TREE4SHOW();
                    //node.tags.Add("工期:" + t.DURATION);
                    node.tags.Add("完成:" + t.FINISH_DATE.Value.ToString("yyyy/MM/dd"));
                    node.tags.Add("開始:" + t.START_DATE.Value.ToString("yyyy/MM/dd"));
                    node.href = "#" + t.PRJ_UID;
                    node.text = t.TASK_NAME;
                    parentnode.addChild(node);
                    //將結點資料記錄至dic 內
                    dicTree.Add(t.PRJ_UID.Value, node);
                    logger.Info("add  node :" + t.PRJ_UID + ",parent=" + t.PRJ_UID);
                }
            }
            return convertToJson(rootnode);
        }
        public string convertToJson(TASK_TREE4SHOW rootnode)
        {
            //將資料集合轉成JSON
            string output = JsonConvert.SerializeObject(rootnode);
            logger.Info("Jason:" + output);
            return output;
        }
        public MapInfoModels getMapView(string projectid,string mapno,string buildno, string primeside,string secondside,string name)
        {
            viewModel.mapDEVICE = getMapItem(projectid, name);
            return viewModel;
        }
        //圖算:設備
        public List<TND_MAP_DEVICE> getMapItem(string projectid,string item_name)
        {
            logger.Info("get map DEVICE info by item_name=" + item_name);
            string sql = "SELECT DEVIVE_ID,M.PROJECT_ID,P.PROJECT_ITEM_ID,MAP_NO,BUILDING_NO "
                + ", M.CREATE_DATE,CREATE_ID,QTY,P.ITEM_DESC LOC_DESC "
                + "FROM TND_MAP_DEVICE M, TND_PROJECT_ITEM P "
                + " WHERE M.PROJECT_ITEM_ID = P.PROJECT_ITEM_ID "
                + " AND P.ITEM_DESC Like @item_name "
                + " AND M.PROJECT_ID = @projectid";
            List<TND_MAP_DEVICE> lstDEVICE = new List<TND_MAP_DEVICE>();
            using (var context = new topmepEntities())
            {
                //條件篩選
                logger.Info(sql);
                lstDEVICE = context.TND_MAP_DEVICE.SqlQuery(sql,
                    new SqlParameter("projectid", projectid),new SqlParameter("item_name","%"+ item_name +"%")).ToList();
            }
            return lstDEVICE;
        }
        //圖算:消防電
        public List<TND_MAP_DEVICE> getMapFP (string projectid,string mapno,string buildno,string primeside, string secondside, string name)
        {
            string sql = "SELECT PROJECT_ID, '消防電-線' AS MAP_TYPE,MAP_NO,BUILDING_NO,PRIMARY_SIDE,SECONDARY_SIDE," 
                + "WIRE_NAME AS PROJECT_ITEM_ID,WIRE_TOTAL_LENGTH AS ITEM_QUANTITY "
                + "FROM TND_MAP_FP AS FP_W "
                + "UNION "
                + "SELECT PROJECT_ID, '消防電-管' AS MAP_TYPE, MAP_NO, BUILDING_NO, PRIMARY_SIDE, SECONDARY_SIDE,"
                + "PIPE_NAME AS PROJECT_ITEM_ID, PIPE_TOTAL_LENGTH AS ITEM_QUANTITY "
                + "FROM TND_MAP_FP AS FP_P";
            return null; 
        }
    }
    #endregion
}