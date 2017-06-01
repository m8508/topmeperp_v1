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
using Newtonsoft.Json;
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
                + " WHERE P.PARENT_UID = B.PRJ_UID and P.TASK_NAME is not null )"
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
                + " WHERE P.PARENT_UID = B.PRJ_UID and P.TASK_NAME is not null )"
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
    #region Tree Stucture
    public abstract class TreeNode<T>
    {
        public T Value { get; set; }
        public abstract TreeNode<T> Parent { get; }
        public abstract TreeList<T> Children { get; }
        public abstract int Count { get; }
        public abstract int Degree { get; }
        public abstract int Depth { get; }
        public abstract int Level { get; }
        public TreeNode(T value)
        {
            this.Value = value;
        }
        public abstract void Add(T value);
        public abstract void Add(TreeNode<T> tree);
        public abstract void Remove();
        public abstract TreeNode<T> Clone();
    }

    public abstract class TreeList<T> : IEnumerable<TreeNode<T>>
    {
        public abstract int Count { get; }
        public abstract IEnumerator<TreeNode<T>> GetEnumerator();

        IEnumerator<TreeNode<T>> IEnumerable<TreeNode<T>>.GetEnumerator()
        {
            return GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    public class LinkedTree<T> : TreeNode<T>
    {
        protected LinkedList<LinkedTree<T>> childrenList;

        protected LinkedTree<T> parent;
        public override TreeNode<T> Parent
        {
            get
            {
                return parent;
            }
        }

        protected LinkedTreeList<T> children;
        public override TreeList<T> Children
        {
            get
            {
                return children;
            }
        }

        public override int Degree
        {
            get
            {
                return childrenList.Count;
            }
        }

        protected int count;
        public override int Count
        {
            get
            {
                return count;
            }
        }

        protected int depth;
        public override int Depth
        {
            get
            {
                return depth;
            }
        }

        protected int level;
        public override int Level
        {
            get
            {
                return level;
            }
        }

        public LinkedTree(T value)
            : base(value)
        {
            childrenList = new LinkedList<LinkedTree<T>>();
            children = new LinkedTreeList<T>(childrenList);
            depth = 1;
            level = 1;
            count = 1;
        }

        public override void Add(T value)
        {
            Add(new LinkedTree<T>(value));
        }

        public override void Add(TreeNode<T> tree)
        {
            LinkedTree<T> gtree = (LinkedTree<T>)tree;
            if (gtree.Parent != null)
                gtree.Remove();
            gtree.parent = this;
            if (gtree.depth + 1 > depth)
            {
                depth = gtree.depth + 1;
                BubbleDepth();
            }
            gtree.level = level + 1;
            gtree.UpdateLevel();
            childrenList.AddLast(gtree);
            count += tree.Count;
            BubbleCount(tree.Count);
        }

        public override void Remove()
        {
            if (parent == null)
                return;
            parent.childrenList.Remove(this);
            if (depth + 1 == parent.depth)
                parent.UpdateDepth();
            parent.count -= count;
            parent.BubbleCount(-count);
            parent = null;
        }

        public override TreeNode<T> Clone()
        {
            return Clone(1);
        }

        protected LinkedTree<T> Clone(int level)
        {
            LinkedTree<T> cloneTree = new LinkedTree<T>(Value);
            cloneTree.depth = depth;
            cloneTree.level = level;
            cloneTree.count = count;
            foreach (LinkedTree<T> child in childrenList)
            {
                LinkedTree<T> cloneChild = child.Clone(level + 1);
                cloneChild.parent = cloneTree;
                cloneTree.childrenList.AddLast(cloneChild);
            }
            return cloneTree;
        }

        protected void BubbleDepth()
        {
            if (parent == null)
                return;

            if (depth + 1 > parent.depth)
            {
                parent.depth = depth + 1;
                parent.BubbleDepth();
            }
        }

        protected void UpdateDepth()
        {
            int tmpDepth = depth;
            depth = 1;
            foreach (LinkedTree<T> child in childrenList)
                if (child.depth + 1 > depth)
                    depth = child.depth + 1;
            if (tmpDepth == depth || parent == null)
                return;
            if (tmpDepth + 1 == parent.depth)
                parent.UpdateDepth();
        }

        protected void BubbleCount(int diff)
        {
            if (parent == null)
                return;

            parent.count += diff;
            parent.BubbleCount(diff);
        }

        protected void UpdateLevel()
        {
            int childLevel = level + 1;
            foreach (LinkedTree<T> child in childrenList)
            {
                child.level = childLevel;
                child.UpdateLevel();
            }
        }
    }

    public class LinkedTreeList<T> : TreeList<T>
    {
        protected LinkedList<LinkedTree<T>> list;

        public LinkedTreeList(LinkedList<LinkedTree<T>> list)
        {
            this.list = list;
        }

        public override int Count
        {
            get
            {
                return list.Count;
            }
        }

        public override IEnumerator<TreeNode<T>> GetEnumerator()
        {
            return list.GetEnumerator();
        }
    }
    #endregion
}
