using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using topmeperp.Models;
using System.Data.Entity.Migrations;

namespace topmeperp.Service
{
    #region 專案進度管理
    public class ProjectPlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string KEY_ID = "RPT";
        //專案任務調用圖算數量物件
        public MapInfoModels viewModel = new MapInfoModels();
        public string resultMessage = "";
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
                if (t.PARENT_UID == 0 || dicTree.Count == 0)
                {
                    //rootnode.tags.Add("工期:" + t.DURATION);
                    rootnode.tags.Add("完成:" + t.FINISH_DATE.Value.ToString("yyyy/MM/dd"));
                    rootnode.tags.Add("開始:" + t.START_DATE.Value.ToString("yyyy/MM/dd"));
                    rootnode.href = t.PRJ_UID.ToString();
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
                    node.href = t.PRJ_UID.ToString();
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
        #region 取得圖算數量功能
        public MapInfoModels getMapView(string projectid, string mapno, string buildno, string primeside, string secondside, string name)
        {
            return viewModel;
        }
        //圖算:設備
        public void getMapItem(string projectid, string item_name, string startid, string endid)
        {
            logger.Info("get map DEVICE info by item_name=" + item_name);
            string sql = "SELECT DEVIVE_ID,M.PROJECT_ID,P.PROJECT_ITEM_ID,MAP_NO,BUILDING_NO "
                + ", M.CREATE_DATE,CREATE_ID,QTY,P.ITEM_DESC LOC_DESC "
                + "FROM TND_MAP_DEVICE M, TND_PROJECT_ITEM P "
                + " WHERE M.PROJECT_ITEM_ID = P.PROJECT_ITEM_ID "
                + " AND P.ITEM_DESC Like @item_name "
                + " AND M.PROJECT_ID = @projectid ";
            //條件篩選
            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            parameters.Add(new SqlParameter("item_name", "%" + item_name + "%"));

            if (null != startid && "" != startid && null != endid && "" != endid)
            {
                sql = sql + " AND CAST(SUBSTRING(P.PROJECT_ITEM_ID,8,LEN(P.PROJECT_ITEM_ID)) as INT) BETWEEN @startid AND @endid ";
                parameters.Add(new SqlParameter("startid", int.Parse(startid)));
                parameters.Add(new SqlParameter("endid", int.Parse(endid)));
            }
            sql = sql + "ORDER BY CAST(SUBSTRING(P.PROJECT_ITEM_ID,8,LEN(P.PROJECT_ITEM_ID)) as INT) ";
            List<TND_MAP_DEVICE> lstDEVICE = new List<TND_MAP_DEVICE>();
            using (var context = new topmepEntities())
            {

                logger.Info("MapItem:" + sql);
                lstDEVICE = context.TND_MAP_DEVICE.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            viewModel.mapDEVICE = lstDEVICE;
            resultMessage = resultMessage + "設備資料筆數:" + lstDEVICE.Count + ",";
        }
        //圖算:消防電
        public void getMapFP(string projectid, string mapno, string buildno, string primeside, string secondside, string name)
        {
            List<TND_PROJECT_ITEM> lstMapFP = null;
            string sql_pipe = "SELECT PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(PIPE_TOTAL_LENGTH) FROM TND_MAP_FP FP WHERE FP.PIPE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN(SELECT PIPE_NAME FROM TND_MAP_FP WHERE TND_MAP_FP.PROJECT_ID=@projectid  ";

            string sql_wire = "SELECT PROJECT_ITEM_ID, PROJECT_ID, ITEM_ID, ITEM_DESC, ITEM_UNIT "
                + ",(SELECT SUM(WIRE_TOTAL_LENGTH) FROM TND_MAP_FP FP WHERE FP.WIRE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN(SELECT WIRE_NAME FROM TND_MAP_FP WHERE TND_MAP_FP.PROJECT_ID=@projectid  ";

            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            if (null != mapno && mapno != "") //圖號
            {
                sql_pipe = sql_pipe + " AND MAP_NO LIKE @mapno";
                sql_wire = sql_wire + " AND MAP_NO LIKE @mapno";
                parameters.Add(new SqlParameter("mapno", "%" + mapno + "%"));
            }
            if (null != buildno && buildno != "")//建築名稱
            {
                sql_pipe = sql_pipe + " AND BUILDING_NO LIKE @buildno";
                sql_wire = sql_wire + " AND BUILDING_NO LIKE @buildno";
                parameters.Add(new SqlParameter("buildno", "%" + buildno + "%"));
            }
            if (null != primeside && primeside != "")//一次側名稱
            {
                sql_pipe = sql_pipe + " AND PRIMARY_SIDE LIKE @primeside";
                sql_wire = sql_wire + " AND PRIMARY_SIDE LIKE @primeside";
                parameters.Add(new SqlParameter("primeside", "%" + primeside + "%"));
            }
            if (null != secondside && secondside != "")//二次側名稱
            {
                sql_pipe = sql_pipe + " AND SECONDARY_SIDE LIKE @secondside";
                sql_wire = sql_wire + " AND SECONDARY_SIDE LIKE @secondside";
                parameters.Add(new SqlParameter("secondside", "%" + secondside + "%"));
            }
            if (null != name && name != "")//品項名稱
            {
                sql_pipe = sql_pipe + " AND P.ITEM_DESC LIKE @name";
                sql_wire = sql_wire + " AND P.ITEM_DESC LIKE @name";
                parameters.Add(new SqlParameter("name", "%" + name + "%"));
            }
            string sql = "SELECT * FROM ((" + sql_pipe + ")) UNION (" + sql_wire + "))) a ORDER BY EXCEL_ROW_ID";
            logger.Debug("PEP SQL=" + sql);
            using (var context = new topmepEntities())
            {
                lstMapFP = context.TND_PROJECT_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            viewModel.ProjectItemInMapFP = lstMapFP;
            resultMessage = resultMessage + "消防電資料筆數:" + lstMapFP.Count + ",";
        }
        //消防水
        public void getMapFW(string projectid, string mapno, string buildno, string primeside, string secondside, string name)
        {
            string sql = "SELECT PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(PIPE_TOTAL_LENGTH)  FROM TND_MAP_FW PLU WHERE PLU.PIPE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID IN (SELECT PIPE_NAME FROM TND_MAP_FW WHERE TND_MAP_FW.PROJECT_ID=@projectid ";

            List<TND_PROJECT_ITEM> lstMapFW = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                if (null != mapno && mapno != "") //圖號
                {
                    sql = sql + " AND MAP_NO LIKE @mapno";
                    parameters.Add(new SqlParameter("mapno", "%" + mapno + "%"));
                }
                if (null != buildno && buildno != "")//建築名稱
                {
                    sql = sql + " AND BUILDING_NO LIKE @buildno";
                    parameters.Add(new SqlParameter("buildno", "%" + buildno + "%"));
                }
                if (null != primeside && primeside != "")//一次側名稱
                {
                    sql = sql + " AND PRIMARY_SIDE LIKE @primeside";
                    parameters.Add(new SqlParameter("primeside", "%" + primeside + "%"));
                }
                if (null != secondside && secondside != "")//二次側名稱
                {
                    sql = sql + " AND SECONDARY_SIDE LIKE @secondside";
                    parameters.Add(new SqlParameter("secondside", "%" + secondside + "%"));
                }
                if (null != name && name != "")//品項名稱
                {
                    sql = sql + " AND P.ITEM_DESC LIKE @name";
                    parameters.Add(new SqlParameter("name", "%" + name + "%"));
                }
                sql = sql + ") ORDER BY EXCEL_ROW_ID";
                logger.Info("MapFW:" + sql);
                lstMapFW = context.TND_PROJECT_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            viewModel.ProjectItemInMapFW = lstMapFW;
            resultMessage = resultMessage + "消防水資料筆數:" + lstMapFW.Count + ",";
        }
        //給排水
        public void getMapPLU(string projectid, string mapno, string buildno, string primeside, string secondside, string name)
        {
            string sql = "SELECT PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                          + ",(SELECT SUM(PIPE_TOTAL_LENGTH) FROM TND_MAP_PLU PLU WHERE PLU.PIPE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                          + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                          + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                          + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                          + "FROM TND_PROJECT_ITEM P "
                          + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID IN (SELECT PIPE_NAME FROM TND_MAP_PLU WHERE TND_MAP_PLU.PROJECT_ID=@projectid ";

            List<TND_PROJECT_ITEM> lstMapPlu = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                if (null != mapno && mapno != "") //圖號
                {
                    sql = sql + " AND MAP_NO LIKE @mapno";
                    parameters.Add(new SqlParameter("mapno", "%" + mapno + "%"));
                }
                if (null != buildno && buildno != "")//建築名稱
                {
                    sql = sql + " AND BUILDING_NO LIKE @buildno";
                    parameters.Add(new SqlParameter("buildno", "%" + buildno + "%"));
                }
                if (null != primeside && primeside != "")//一次側名稱
                {
                    sql = sql + " AND PRIMARY_SIDE LIKE @primeside";
                    parameters.Add(new SqlParameter("primeside", "%" + primeside + "%"));
                }
                if (null != secondside && secondside != "")//二次側名稱
                {
                    sql = sql + " AND SECONDARY_SIDE LIKE @secondside";
                    parameters.Add(new SqlParameter("secondside", "%" + secondside + "%"));
                }
                if (null != name && name != "")//品項名稱
                {
                    sql = sql + " AND P.ITEM_DESC LIKE @name";
                    parameters.Add(new SqlParameter("name", "%" + name + "%"));
                }
                sql = sql + ") ORDER BY EXCEL_ROW_ID";
                logger.Info("getMapPLU:" + sql);
                lstMapPlu = context.TND_PROJECT_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            viewModel.ProjectItemInMapPLU = lstMapPlu;
            resultMessage = resultMessage + "給排水資料筆數:" + lstMapPlu.Count + ",";
        }
        //電器管線
        public void getMapPEP(string projectid, string mapno, string buildno, string primeside, string secondside, string name)
        {

            List<TND_PROJECT_ITEM> lstMapPEP = null;
            string sql_pipe = "SELECT PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(PIPE_TOTAL_LENGTH) FROM TND_MAP_PEP PEP WHERE PEP.PIPE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN(SELECT PIPE_NAME FROM TND_MAP_PEP WHERE TND_MAP_PEP.PROJECT_ID=@projectid  ";

            string sql_ground = "SELECT PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(GROUND_WIRE_TOTAL_LENGTH) FROM TND_MAP_PEP PEP WHERE PEP.GROUND_WIRE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid  AND P.PROJECT_ITEM_ID "
                + "IN(SELECT GROUND_WIRE_NAME FROM TND_MAP_PEP WHERE TND_MAP_PEP.PROJECT_ID=@projectid ";

            string sql_wire = "SELECT PROJECT_ITEM_ID, PROJECT_ID, ITEM_ID, ITEM_DESC, ITEM_UNIT "
                + ",(SELECT SUM(WIRE_TOTAL_LENGTH) FROM TND_MAP_PEP PEP WHERE PEP.WIRE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN(SELECT WIRE_NAME FROM TND_MAP_PEP WHERE TND_MAP_PEP.PROJECT_ID=@projectid  ";


            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            if (null != mapno && mapno != "") //圖號
            {
                sql_pipe = sql_pipe + " AND MAP_NO LIKE @mapno";
                sql_ground = sql_ground + " AND MAP_NO LIKE @mapno";
                sql_wire = sql_wire + " AND MAP_NO LIKE @mapno";
                parameters.Add(new SqlParameter("mapno", "%" + mapno + "%"));
            }
            if (null != buildno && buildno != "")//建築名稱
            {
                sql_pipe = sql_pipe + " AND BUILDING_NO LIKE @buildno";
                sql_ground = sql_ground + " AND BUILDING_NO LIKE @buildno";
                sql_wire = sql_wire + " AND BUILDING_NO LIKE @buildno";
                parameters.Add(new SqlParameter("buildno", "%" + buildno + "%"));
            }
            if (null != primeside && primeside != "")//一次側名稱
            {
                sql_pipe = sql_pipe + " AND PRIMARY_SIDE LIKE @primeside";
                sql_ground = sql_ground + " AND PRIMARY_SIDE LIKE @primeside";
                sql_wire = sql_wire + " AND PRIMARY_SIDE LIKE @primeside";
                parameters.Add(new SqlParameter("primeside", "%" + primeside + "%"));
            }
            if (null != secondside && secondside != "")//二次側名稱
            {
                sql_pipe = sql_pipe + " AND SECONDARY_SIDE LIKE @secondside";
                sql_ground = sql_ground + " AND SECONDARY_SIDE LIKE @secondside";
                sql_wire = sql_wire + " AND SECONDARY_SIDE LIKE @secondside";
                parameters.Add(new SqlParameter("secondside", "%" + secondside + "%"));
            }
            if (null != name && name != "")//品項名稱
            {
                sql_pipe = sql_pipe + " AND P.ITEM_DESC LIKE @name";
                sql_ground = sql_ground + " AND P.ITEM_DESC LIKE @name";
                sql_wire = sql_wire + " AND P.ITEM_DESC LIKE @name";
                parameters.Add(new SqlParameter("name", "%" + name + "%"));
            }
            string sql = "SELECT * FROM ((" + sql_pipe + ")) UNION (" + sql_ground + ")) UNION (" + sql_wire + "))) a ORDER BY EXCEL_ROW_ID";
            logger.Debug("PEP SQL=" + sql);
            using (var context = new topmepEntities())
            {
                lstMapPEP = context.TND_PROJECT_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            viewModel.ProjectItemInMapPEP = lstMapPEP;
            resultMessage = resultMessage + "電器資料筆數:" + lstMapPEP.Count + ",";
        }
        //弱電
        public void getMapLCP(string projectid, string mapno, string buildno, string primeside, string secondside, string name)
        {
            //主鍵值+ _p1 & _p2 區隔資料
            List<TND_PROJECT_ITEM> lstMapLCP = null;
            string sql_pipe1 = "SELECT PROJECT_ITEM_ID +'_p1' PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(PIPE_1_TOTAL_LEN) FROM TND_MAP_LCP LCP WHERE LCP.PIPE_1_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN (SELECT PIPE_1_NAME FROM TND_MAP_LCP WHERE TND_MAP_LCP.PROJECT_ID=@projectid ";

            string sql_pipe2 = "SELECT PROJECT_ITEM_ID+'_p2' PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(PIPE_2_TOTAL_LEN) FROM TND_MAP_LCP LCP WHERE LCP.PIPE_2_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN (SELECT PIPE_2_NAME FROM TND_MAP_LCP WHERE TND_MAP_LCP.PROJECT_ID=@projectid  ";

            string sql_ground = "SELECT PROJECT_ITEM_ID,PROJECT_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT "
                + ",(SELECT SUM(GROUND_WIRE_TOTAL_LENGTH) FROM TND_MAP_LCP LCP WHERE LCP.GROUND_WIRE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid  AND P.PROJECT_ITEM_ID "
                + "IN(SELECT GROUND_WIRE_NAME FROM TND_MAP_LCP WHERE TND_MAP_LCP.PROJECT_ID=@projectid ";

            string sql_wire = "SELECT PROJECT_ITEM_ID, PROJECT_ID, ITEM_ID, ITEM_DESC, ITEM_UNIT "
                + ",(SELECT SUM(WIRE_TOTAL_LENGTH) FROM TND_MAP_LCP LCP WHERE LCP.WIRE_NAME = P.PROJECT_ITEM_ID) as ITEM_QUANTITY "
                + ",ITEM_UNIT_PRICE,MAN_PRICE,ITEM_REMARK,TYPE_CODE_1,TYPE_CODE_2,SUB_TYPE_CODE,SYSTEM_MAIN ,SYSTEM_SUB "
                + ",MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + ",SHEET_NAME ,EXCEL_ROW_ID,QUO_PRICE "
                + "FROM TND_PROJECT_ITEM P "
                + "WHERE P.PROJECT_ID=@projectid AND P.PROJECT_ITEM_ID "
                + "IN(SELECT WIRE_NAME FROM TND_MAP_LCP WHERE TND_MAP_LCP.PROJECT_ID=@projectid  ";


            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            if (null != mapno && mapno != "") //圖號
            {
                sql_pipe1 = sql_pipe1 + " AND MAP_NO LIKE @mapno";
                sql_pipe2 = sql_pipe2 + " AND MAP_NO LIKE @mapno";
                sql_ground = sql_ground + " AND MAP_NO LIKE @mapno";
                sql_wire = sql_wire + " AND MAP_NO LIKE @mapno";
                parameters.Add(new SqlParameter("mapno", "%" + mapno + "%"));
            }
            if (null != buildno && buildno != "")//建築名稱
            {
                sql_pipe1 = sql_pipe1 + " AND BUILDING_NO LIKE @buildno";
                sql_pipe2 = sql_pipe2 + " AND BUILDING_NO LIKE @buildno";
                sql_ground = sql_ground + " AND BUILDING_NO LIKE @buildno";
                sql_wire = sql_wire + " AND BUILDING_NO LIKE @buildno";
                parameters.Add(new SqlParameter("buildno", "%" + buildno + "%"));
            }
            if (null != primeside && primeside != "")//一次側名稱
            {
                sql_pipe1 = sql_pipe1 + " AND PRIMARY_SIDE LIKE @primeside";
                sql_pipe2 = sql_pipe2 + " AND PRIMARY_SIDE LIKE @primeside";
                sql_ground = sql_ground + " AND PRIMARY_SIDE LIKE @primeside";
                sql_wire = sql_wire + " AND PRIMARY_SIDE LIKE @primeside";
                parameters.Add(new SqlParameter("primeside", "%" + primeside + "%"));
            }
            if (null != secondside && secondside != "")//二次側名稱
            {
                sql_pipe1 = sql_pipe1 + " AND SECONDARY_SIDE LIKE @secondside";
                sql_pipe2 = sql_pipe2 + " AND SECONDARY_SIDE LIKE @secondside";
                sql_ground = sql_ground + " AND SECONDARY_SIDE LIKE @secondside";
                sql_wire = sql_wire + " AND SECONDARY_SIDE LIKE @secondside";
                parameters.Add(new SqlParameter("secondside", "%" + secondside + "%"));
            }
            if (null != name && name != "")//品項名稱
            {
                sql_pipe1 = sql_pipe1 + " AND P.ITEM_DESC LIKE @name";
                sql_pipe2 = sql_pipe2 + " AND P.ITEM_DESC LIKE @name";
                sql_ground = sql_ground + " AND P.ITEM_DESC LIKE @name";
                sql_wire = sql_wire + " AND P.ITEM_DESC LIKE @name";
                parameters.Add(new SqlParameter("name", "%" + name + "%"));
            }
            string sql = "SELECT * FROM ((" + sql_pipe1 + ")) UNION (" + sql_pipe2 + ")) UNION (" + sql_ground + ")) UNION (" + sql_wire + "))) a ORDER BY EXCEL_ROW_ID";
            logger.Debug("LCP SQL=" + sql);
            using (var context = new topmepEntities())
            {
                lstMapLCP = context.TND_PROJECT_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            for (int i = 0; i < lstMapLCP.Count; i++)
            {
                logger.Debug("item id=" + lstMapLCP[i].ITEM_DESC + ",Qty=" + lstMapLCP[i].ITEM_QUANTITY);
            }
            viewModel.ProjectItemInMapLCP = lstMapLCP;
            resultMessage = resultMessage + "弱電資料筆數:" + lstMapLCP.Count + ",";
        }
        #endregion
        #region //設定任務與圖算項目
        //設備
        public int choiceMapItem(string projectid, string prjuid, string mapdeviceIds)
        {
            int i = -1;
            logger.Info("projectid=" + projectid + ",prjuid=" + prjuid + ",mapdeviceIds=" + mapdeviceIds);

            using (var context = new topmepEntities())
            {
                //清除原來任務之工作項目，再將設備資料寫入Task2MapItem
                string sql = "DELETE PLAN_TASK2MAPITEM WHERE PROJECT_ID=@projectId AND PRJ_UID=@prjuid AND　MAP_TYPE='TND_MAP_DEVICE';"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_DEVICE' AS MAP_TYPE, DEVIVE_ID AS MAP_PK, PROJECT_ITEM_ID  FROM TND_MAP_DEVICE "
                    + " WHERE DEVIVE_ID in (" + @mapdeviceIds + ");";
                logger.Debug(sql);
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        //設定任務與圖算項目-電器管線
        public int choiceMapItemPEP(string projectid, string prjuid, string mapPepIds)
        {
            int i = -1;
            logger.Info("projectid=" + projectid + ",prjuid=" + prjuid + ",mapPepIds=" + mapPepIds);

            using (var context = new topmepEntities())
            {
                mapPepIds = mapPepIds.Replace(",", "','");
                //清除原來任務之工作項目，再將設備資料寫入Task2MapItem
                string sql = "DELETE PLAN_TASK2MAPITEM WHERE PROJECT_ID=@projectId AND PRJ_UID=@prjuid AND　MAP_TYPE='TND_MAP_PEP';"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_PEP' AS MAP_TYPE, 0 AS MAP_PK, PIPE_NAME  FROM TND_MAP_PEP "
                    + " WHERE PIPE_NAME in ('" + mapPepIds + "');"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_PEP' AS MAP_TYPE, 1 AS MAP_PK, WIRE_NAME  FROM TND_MAP_PEP "
                    + " WHERE WIRE_NAME in ('" + mapPepIds + "');"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_PEP' AS MAP_TYPE, 2 AS MAP_PK, GROUND_WIRE_NAME  FROM TND_MAP_PEP "
                    + " WHERE GROUND_WIRE_NAME in ('" + mapPepIds + "');";
                logger.Debug(sql);
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        //設定任務與圖算項目-弱電
        public int choiceMapItemLCP(string projectid, string prjuid, string mapLcpIds)
        {
            int i = -1;
            logger.Info("projectid=" + projectid + ",prjuid=" + prjuid + ",mapLcpIds=" + mapLcpIds);

            using (var context = new topmepEntities())
            {
                mapLcpIds = mapLcpIds.Replace("_p1", "");
                mapLcpIds = mapLcpIds.Replace("_p2", "");
                mapLcpIds = mapLcpIds.Replace(",", "','");

                //清除原來任務之工作項目，再將設備資料寫入Task2MapItem
                string sql = "DELETE PLAN_TASK2MAPITEM WHERE PROJECT_ID=@projectId AND PRJ_UID=@prjuid AND　MAP_TYPE='TND_MAP_LCP';"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_LCP' AS MAP_TYPE, 0 AS MAP_PK, PIPE_1_NAME  FROM TND_MAP_LCP "
                    + " WHERE PIPE_1_NAME in ('" + mapLcpIds + "');"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_LCP' AS MAP_TYPE, 1 AS MAP_PK, PIPE_2_NAME  FROM TND_MAP_LCP "
                    + " WHERE PIPE_2_NAME in ('" + mapLcpIds + "');"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_LCP' AS MAP_TYPE, 2 AS MAP_PK, WIRE_NAME  FROM TND_MAP_LCP "
                    + " WHERE WIRE_NAME in ('" + mapLcpIds + "');"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_LCP' AS MAP_TYPE, 3 AS MAP_PK, GROUND_WIRE_NAME  FROM TND_MAP_LCP "
                    + " WHERE GROUND_WIRE_NAME in ('" + mapLcpIds + "');";
                logger.Debug(sql);
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        //設定任務與圖算項目-消防電
        public int choiceMapItemFP(string projectid, string prjuid, string mapFpIds)
        {
            int i = -1;
            logger.Info("projectid=" + projectid + ",prjuid=" + prjuid + ",mapFpIds=" + mapFpIds);

            using (var context = new topmepEntities())
            {
                mapFpIds = mapFpIds.Replace(",", "','");
                //清除原來任務之工作項目，再將設備資料寫入Task2MapItem
                string sql = "DELETE PLAN_TASK2MAPITEM WHERE PROJECT_ID=@projectId AND PRJ_UID=@prjuid AND　MAP_TYPE='TND_MAP_FP';"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_FP' AS MAP_TYPE, 0 AS MAP_PK, PIPE_NAME  FROM TND_MAP_FP "
                    + " WHERE PIPE_NAME in ('" + mapFpIds + "');"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_FP' AS MAP_TYPE, 1 AS MAP_PK, WIRE_NAME  FROM TND_MAP_FP "
                    + " WHERE WIRE_NAME in ('" + mapFpIds + "');";
                logger.Debug(sql);
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        //設定任務與圖算項目-消防水
        public int choiceMapItemFW(string projectid, string prjuid, string mapPitemIds)
        {
            int i = -1;
            logger.Info("projectid=" + projectid + ",prjuid=" + prjuid + ",mapPitemIds=" + mapPitemIds);

            using (var context = new topmepEntities())
            {
                mapPitemIds = mapPitemIds.Replace(",", "','");
                //清除原來任務之工作項目，再將設備資料寫入Task2MapItem
                string sql = "DELETE PLAN_TASK2MAPITEM WHERE PROJECT_ID=@projectId AND PRJ_UID=@prjuid AND　MAP_TYPE='TND_MAP_FW';"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + " SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid AS PRJ_UID,'TND_MAP_FW' AS MAP_TYPE, 0 AS MAP_PK, PIPE_NAME  FROM TND_MAP_FW "
                    + " WHERE PIPE_NAME in ('" + mapPitemIds + "');";
                logger.Debug(sql);
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        //設定任務與圖算項目-給排水
        public int choiceMapItemPLU(string projectid, string prjuid, string mappluIds)
        {
            int i = -1;
            logger.Info("projectid=" + projectid + ",prjuid=" + prjuid + ",MAP_PLU=" + mappluIds);

            using (var context = new topmepEntities())
            {
                mappluIds = mappluIds.Replace(",", "','");
                //清除原來任務之工作項目，再將設備資料寫入Task2MapItem
                string sql = "DELETE PLAN_TASK2MAPITEM WHERE PROJECT_ID=@projectId AND PRJ_UID=@prjuid AND　MAP_TYPE='TND_MAP_PLU';"
                    + "INSERT INTO PLAN_TASK2MAPITEM (PROJECT_ID,PRJ_UID,MAP_TYPE,MAP_PK,PROJECT_ITEM_ID) "
                    + "SELECT DISTINCT @projectId AS PROJECT_ID,@prjuid  AS PRJ_UID,'TND_MAP_PLU' AS MAP_TYPE, 0 AS MAP_PK, PIPE_NAME  FROM TND_MAP_PLU "
                    + " WHERE PIPE_NAME in ('" + mappluIds + "');";
                logger.Debug(sql);
                var parameters = new List<SqlParameter>();
                //設定專案名編號資料
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("prjuid", prjuid));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        #endregion
        public List<TND_PROJECT_ITEM> getItemInTask(string projectid, string prjuid)
        {
            logger.Info("get ItemTask Project_id=" + projectid + ",prjuid=" + prjuid);
            List<TND_PROJECT_ITEM> lstProjectItem = new List<TND_PROJECT_ITEM>();
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = "SELECT * FROM TND_PROJECT_ITEM WHERE PROJECT_ITEM_ID IN (SELECT PROJECT_ITEM_ID FROM PLAN_TASK2MAPITEM "
                        + "WHERE PROJECT_ID=@projectid AND PRJ_UID=@prjuid);";
                    logger.Debug("sql=" + sql);
                    lstProjectItem = context.TND_PROJECT_ITEM.SqlQuery(sql, new SqlParameter("projectid", projectid), new SqlParameter("prjuid", int.Parse(prjuid))).ToList();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
            }
            return lstProjectItem;
        }
        //取得特定日期專案任務清單
        public List<PLAN_TASK> getTaskByDate(string projectid, DateTime dt)
        {
            List<PLAN_TASK> lstTask = new List<PLAN_TASK>();
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = "SELECT [TASK_ID],[PROJECT_ID],[PRJ_ID],[PRJ_UID],[TASK_NAME],[START_DATE],[FINISH_DATE] "
                             + ",[PARENT_UID], CONVERT(varchar,DATEDIFF (day, START_DATE, FINISH_DATE)) as DURATION,[ROOT_TAG] "
                             + ",[CREATE_ID],[CREATE_DATE],[MODIFY_ID],[MODIFY_DATE]"
                             + " FROM PLAN_TASK WHERE PROJECT_ID=@projectid  AND CONVERT(datetime, @dt, 20)  BETWEEN START_DATE AND FINISH_DATE "
                             + " AND PRJ_ID >= (SELECT PRJ_ID FROM PLAN_TASK WHERE PROJECT_ID=@projectid  AND ROOT_TAG = 'Y') "
                             + " ORDER BY DATEDIFF(day, START_DATE, FINISH_DATE);";
                    logger.Debug("sql=" + sql);
                    logger.Debug("dt" + dt.ToString("yyyy-MM-dd"));
                    lstTask = context.PLAN_TASK.SqlQuery(sql, new SqlParameter("projectid", projectid), new SqlParameter("dt", @dt)).ToList();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
            }
            logger.Info("get task by now:" + lstTask.Count);
            return lstTask;
        }
        //讀取專案基本資料
        public TND_PROJECT getProject(string projectid)
        {
            TnderProject service = new TnderProject();
            return service.getProjectById(projectid);
        }
        //讀取任務基本資料
        public PLAN_TASK getProjectTask(string projectid, int prjuid)
        {
            string sql = "SELECT * FROM PLAN_TASK WHERE PROJECT_ID = @projectid and PRJ_UID=@prjuid;";
            PLAN_TASK planTask = null;
            using (var context = new topmepEntities())
            {
                planTask = context.PLAN_TASK.SqlQuery(sql, new SqlParameter("projectid", projectid), new SqlParameter("prjuid", prjuid)).First();
                logger.Debug("Task Info=" + planTask.TASK_NAME);
            }
            return planTask;
        }
        //建立新的施工日報資料使用
        public DailyReport newDailyReport(string projectid, int prjuid)
        {
            DailyReport newDailyRpt = new DailyReport();
            //建立料件資料
            newDailyRpt.lstDailyRptItem4Show = getItem(projectid, prjuid);
            //新報告上無Report ID 用假資料
            newDailyRpt.lstDailyRptWokerType4Show = getDailyReportRecord4Worker(projectid,"000", "Worker"); //SystemParameter.getSystemPara("ProjectPlanService", "Worker");
            newDailyRpt.lstDailyRptMachine4Show = getDailyReportRecord4Worker(projectid,"000", "Machine"); //SystemParameter.getSystemPara("ProjectPlanService", "Machine");
            return newDailyRpt;
        }
        public DailyReport getDailyReport(string reportId)
        {
            DailyReport drDailyRpt = new DailyReport();
            using (var context = new topmepEntities())
            {
                string sql = "SELECT * FROM PLAN_DALIY_REPORT WHERE REPORT_ID=@reportId ";
                logger.Debug("get daily report ,sql=" + sql + ",reportId=" + reportId);
                drDailyRpt.dailyRpt = context.PLAN_DALIY_REPORT.SqlQuery(sql, new SqlParameter("reportId", reportId)).First();
                drDailyRpt.lstDailyRptItem4Show = getItem(reportId);
                drDailyRpt.lstRptTask = getTaskByReportId(reportId);
                drDailyRpt.lstDailyRptWokerType4Show = getDailyReportRecord4Worker(drDailyRpt.dailyRpt.PROJECT_ID, reportId, "Worker");
                drDailyRpt.lstDailyRptMachine4Show = getDailyReportRecord4Worker(drDailyRpt.dailyRpt.PROJECT_ID, reportId, "Machine");
                //取得重要事件
                sql = "SELECT * FROM PLAN_DR_NOTE WHERE REPORT_ID=@reportId";
                logger.Debug("get notes ,sql=" + sql + ",reportId=" + reportId);
                drDailyRpt.lstRptNote = context.PLAN_DR_NOTE.SqlQuery(sql, new SqlParameter("reportId", reportId)).ToList() ;
            }
            return drDailyRpt;
        }

        //建立日報相關紀錄
        public string createDailyReport(DailyReport dr)
        {
            try
            {
                using (var db = new topmepEntities())
                {
                    //1.將日報相關資料刪除
                    string sql = "DELETE FROM PLAN_DR_ITEM WHERE REPORT_ID=@reportId;DELETE FROM PLAN_DR_NOTE WHERE REPORT_ID = @reportId;DELETE FROM PLAN_DR_WORKER WHERE REPORT_ID = @reportId;DELETE FROM PLAN_DR_TASK WHERE REPORT_ID = @reportId;";
                    logger.Info("remove daily Report:" + sql + ",reportid="+ dr.dailyRpt.REPORT_ID);
                    int i = db.Database.ExecuteSqlCommand(sql, new SqlParameter("reportId", dr.dailyRpt.REPORT_ID));
                    logger.Debug("remove daily Report:" + i);
                    //2.
                    db.PLAN_DALIY_REPORT.AddOrUpdate(dr.dailyRpt);
                    db.PLAN_DR_TASK.AddRange(dr.lstRptTask);
                    db.PLAN_DR_ITEM.AddRange(dr.lstRptItem);
                    db.PLAN_DR_WORKER.AddRange(dr.lstRptWorkerAndMachine);
                    db.PLAN_DR_NOTE.AddRange(dr.lstRptNote);
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
                return "新增日報失敗(" + ex.Message + ")";
            }

            return "新增日報成功!!";
        }
        //取得專案任務對應之料件項目
        public List<DailyReportItem> getItem(string projectid, int prjuid)
        {
            List<DailyReportItem> lstDailyRptItem = new List<DailyReportItem>();
            string sql = "SELECT TI.TASKUID,TI.PRJ_UID,TI.PROJECT_ID,TI.PROJECT_ITEM_ID,"
                + " (SELECT ITEM_DESC FROM TND_PROJECT_ITEM i WHERE i.PROJECT_ITEM_ID=TI.PROJECT_ITEM_ID) ITEM_DESC,"
                + " (SELECT SUM(ISNULL(FINISH_QTY,0)) FROM PLAN_DR_ITEM WHERE PLAN_ITEM_ID = TI.PROJECT_ITEM_ID) as ACCUMULATE_QTY ,QTY,null AS FINSIH_QTY "
                + " FROM PLAN_TASK2MAPITEM TI LEFT OUTER JOIN vw_MAP_MATERLIALIST_DETAIL MAP "
                + " ON TI.PROJECT_ITEM_ID = MAP.PROJECT_ITEM_ID "
                + " WHERE TI.PROJECT_ID = @projectid AND TI.PRJ_UID = @prjuid "
                + "ORDER BY  CAST(SUBSTRING(TI.PROJECT_ITEM_ID, 8, LEN(TI.PROJECT_ITEM_ID)) as INT);";
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql + ",projectid=" + projectid + ",prjuid=" + prjuid);
                lstDailyRptItem = context.Database.SqlQuery<DailyReportItem>(sql, new SqlParameter("projectid", projectid), new SqlParameter("prjuid",prjuid)).ToList();
                logger.Debug("lstDailyRptItem count=" + lstDailyRptItem.Count);
            }
            return lstDailyRptItem;
        }
        public List<DailyReportItem> getItem(string reportId)
        {
            List<DailyReportItem> lstDailyRptItem = new List<DailyReportItem>();
            string sql = "SELECT DR_ITEM_ID AS TASKUID,0 as PRJ_UID,i.PROJECT_ID,PLAN_ITEM_ID as PROJECT_ITEM_ID,QTY,"
                +"(SELECT ITEM_DESC FROM TND_PROJECT_ITEM p WHERE i.PLAN_ITEM_ID = p.PROJECT_ITEM_ID) AS ITEM_DESC, LAST_QTY AS ACCUMULATE_QTY, FINISH_QTY "
                + "FROM PLAN_DR_ITEM i,vw_MAP_MATERLIALIST_DETAIL Map WHERE REPORT_ID = @reportId AND Map.PROJECT_ITEM_ID=i.PLAN_ITEM_ID; ";
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql + ",reportId=" + reportId );
                lstDailyRptItem = context.Database.SqlQuery<DailyReportItem>(sql, new SqlParameter("reportId", reportId)).ToList();
                logger.Debug("lstDailyRptItem count=" + lstDailyRptItem.Count);
            }
            return lstDailyRptItem;
        }
        //取得人工/機具使用數量
        public List<DailyReportRecord4Worker> getDailyReportRecord4Worker(string projectid,string reportid, string type)
        {
            //string sql = "SELECT PA.FUNCTION_ID,PA.KEY_FIELD,PA.VALUE_FIELD,"
            //    + "(SELECT SUM(ISNULL(WORKER_QTY, 0)) AS LAST_QTY FROM PLAN_DR_WORKER WHERE REPORT_ID = @ReportID AND PARA_KEY_ID = RPT.PARA_KEY_ID) AS LAST_QTY,"
            //    + "RPT.WORKER_QTY,RPT.REMARK,RPT.REPORT_ID "
            //    + "FROM  SYS_PARA PA LEFT OUTER JOIN(SELECT * FROM PLAN_DR_WORKER WHERE REPORT_ID = @ReportID ) RPT "
            //    + " ON PA.KEY_FIELD = RPT.PARA_KEY_ID WHERE FUNCTION_ID = 'ProjectPlanService' AND FIELD_ID = @type ORDER BY KEY_FIELD; ";

            string sql = "SELECT PA.FUNCTION_ID,PA.KEY_FIELD,PA.VALUE_FIELD,"
                + "(SELECT SUM(ISNULL(WORKER_QTY, 0)) AS LAST_QTY FROM PLAN_DR_WORKER, PLAN_DALIY_REPORT "
                + "WHERE PLAN_DALIY_REPORT.REPORT_ID = PLAN_DR_WORKER.REPORT_ID "
                + "AND PLAN_DALIY_REPORT.PROJECT_ID = @projectid AND PLAN_DR_WORKER.PARA_KEY_ID = PA.KEY_FIELD) AS LAST_QTY, "
                + "RPT.WORKER_QTY,RPT.REMARK,RPT.REPORT_ID FROM  SYS_PARA PA LEFT OUTER JOIN(SELECT * FROM PLAN_DR_WORKER WHERE REPORT_ID = @reportid) RPT ON PA.KEY_FIELD = RPT.PARA_KEY_ID "
                + "WHERE FUNCTION_ID = 'ProjectPlanService' AND FIELD_ID =@type  ORDER BY KEY_FIELD; ";

            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            parameters.Add(new SqlParameter("reportid", reportid));
            parameters.Add(new SqlParameter("type", type));

            List<DailyReportRecord4Worker> lst = new List<DailyReportRecord4Worker>();
            using (var context = new topmepEntities())
            {
                logger.Info("get DailyReportRecord4Worker sql:" + sql +",type=" + type);
                lst = context.Database.SqlQuery<DailyReportRecord4Worker>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //取得日報任務資料
        public List<PLAN_DR_TASK> getTaskByReportId(string reportid)
        {
            string sql = "SELECT DT.REPORT_ID,DR_TASK_ID,DT.PROJECT_ID,DT.PRJ_UID,T.TASK_NAME "
                + "FROM PLAN_DR_TASK DT ,PLAN_DALIY_REPORT RPT, PLAN_TASK T "
                + "WHERE DT.REPORT_ID = RPT.REPORT_ID "
                + "AND DT.PROJECT_ID = T.PROJECT_ID AND DT.PRJ_UID = T.PRJ_UID "
                + "AND RPT.REPORT_ID =@ReportID ";
            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("ReportID", reportid));
            List<PLAN_DR_TASK> lst = new List<PLAN_DR_TASK>();
            using (var context = new topmepEntities())
            {
                logger.Info("get task for report sql:" + sql);
                lst = context.Database.SqlQuery<PLAN_DR_TASK>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //取得日報資料，缺乏加總數字
        public List<PLAN_DALIY_REPORT> getDailyReportList(string projectid, DateTime dtStart, DateTime dtEnd, string strSummary)
        {
            List<PLAN_DALIY_REPORT> lst = null;
            string sql = "SELECT REPORT_ID,PROJECT_ID,REPORT_DATE,WEATHER,SUMMARY,SCENE_USER_NAME,SUPERVISION_NAME,OWNER_NAME,"
                + "MODIFY_USER_ID,MODIFY_DATE,CREATE_USER_ID,CREATE_DATE "
                + "FROM PLAN_DALIY_REPORT WHERE PROJECT_ID=@projectid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));

            if (null != dtStart && dtStart != DateTime.MinValue)
            {
                sql = sql + " AND REPORT_DATE BETWEEN @dtStart AND @dtEnd";
                parameters.Add(new SqlParameter("dtStart", dtStart));
                if (null == dtEnd)
                {
                    dtEnd = dtStart.AddSeconds(86399.0);
                }
                parameters.Add(new SqlParameter("dtEnd", dtEnd));
            }
            if (null != strSummary && "" != strSummary)
            {
                sql = sql + " AND SUMMARY LIKE @strSummary";
                parameters.Add(new SqlParameter("strSummary", "'%" + strSummary + "%'"));
            }
            sql = sql + " ORDER BY REPORT_DATE DESC";
            logger.Info("sql=" + sql);
            using (var context = new topmepEntities())
            {
                lst = context.PLAN_DALIY_REPORT.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
    }
    #endregion
}