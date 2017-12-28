using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    //成本預算管制表Service Layer
    public class ContextService4PlanCost : PlanService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public CostControlInfo CostInfo = new CostControlInfo();
        public ContextService4PlanCost()
        {

        }
        public void getCostControlInfo(string projectId)
        {
            logger.Debug("get Cost Controll Info By ProjectId=" + projectId);
            //專案基本資料
            CostInfo.Project = getProject(projectId);
            //1.合約金額與追加減項目
            CostInfo.Revenue = getPlanRevenueById(projectId);
            //1.1 異動單彙整資料
            CostInfo.lstCostChangeEvent = getCostChangeEvnet(projectId);
            //2.直接成本:材料與工資
            PurchaseFormService pfservice = new PurchaseFormService();
            CostInfo.lstDirectCostItem = pfservice.getPlanContract(projectId);
            CostInfo.lstDirectCostItem.AddRange(pfservice.getPlanContract4Wage(projectId));
            //3.間接成本
            CostInfo.lstIndirectCostItem = getIndirectCost();
        }
        //建立間接成本資料
        public void createIndirectCost(string projectId, string userid)
        {
            List<SYS_PARA> lstItem = SystemParameter.getSystemPara("IndirectCostItem");
            List<PLAN_INDIRECT_COST> lstIndirectCostItem = new List<PLAN_INDIRECT_COST>();
            //取得合約金額
            CostInfo.Revenue = getPlanRevenueById(projectId);
            //取得間接成本項目
            foreach (SYS_PARA p in lstItem)
            {
                PLAN_INDIRECT_COST it = new PLAN_INDIRECT_COST();
                it.PROJECT_ID = projectId;
                it.FIELD_ID = p.FIELD_ID;
                it.FIELD_DESC = p.VALUE_FIELD;
                it.PERCENTAGE = decimal.Parse(p.KEY_FIELD);
                it.MODIFY_ID = userid;
                it.MODIFY_DATE = DateTime.Now;
                // System.Convert.ToDoublSystem.Math.Round(1.235, 2, MidpointRounding.AwayFromZero)
                it.COST = Convert.ToDecimal(Math.Round(Convert.ToDouble(CostInfo.Revenue.PLAN_REVENUE * decimal.Parse(p.KEY_FIELD)), 0, MidpointRounding.AwayFromZero));
                logger.Debug(p.VALUE_FIELD + " Indirect Cost=" + it.COST + ",per=" + p.KEY_FIELD);
                lstIndirectCostItem.Add(it);
            }
            using (var context = new topmepEntities())
            {
                ///刪除現有資料
                string sql = "DELETE FROM PLAN_INDIRECT_COST WHERE PROJECT_ID=@projectId";
                logger.Debug("sql=" + sql + ",projectid=" + projectId);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", projectId));
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                logger.Debug("sql=" + sql + ",projectid=" + projectId);
                ///將新資料存入
                context.PLAN_INDIRECT_COST.AddRange(lstIndirectCostItem);
                context.SaveChanges();
            }
        }
        //取得間接成本資料
        private List<PLAN_INDIRECT_COST> getIndirectCost()
        {
            List<PLAN_INDIRECT_COST> lst = null;
            using (var context = new topmepEntities())
            {
                string sql = "SELECT * FROM PLAN_INDIRECT_COST WHERE PROJECT_ID=@projectId ORDER BY FIELD_ID";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", CostInfo.Project.PROJECT_ID));
                logger.Debug("SQL:" + sql + ",projectId=" + CostInfo.Project.PROJECT_ID);
                lst = context.PLAN_INDIRECT_COST.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //修正間接成本資料
        public void modifyIndirectCost(string projectId, List<PLAN_INDIRECT_COST> items)
        {
            using (var context = new topmepEntities())
            {
                ///逐筆更新資料
                string sql = "UPDATE PLAN_INDIRECT_COST SET COST = @cost, MODIFY_ID = @modifyId, MODIFY_DATE = @modifyDate, NOTE = ISNULL(Note,'') + @Note  WHERE PROJECT_ID = @projectId AND FIELD_ID = @fieldId";
                logger.Debug("sql=" + sql + ",projectid=" + projectId);
                foreach (PLAN_INDIRECT_COST it in items)
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("projectId", projectId));
                    parameters.Add(new SqlParameter("fieldId", it.FIELD_ID));
                    parameters.Add(new SqlParameter("cost", it.COST));
                    parameters.Add(new SqlParameter("modifyId", it.MODIFY_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("Note", it.NOTE));
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                    logger.Debug("sql=" + sql + ",projectid=" + projectId);
                }
                ///將新資料存入
                context.SaveChanges();
            }
        }

        ///成本異動彙整資料(對業主)
        public List<CostChangeEvent> getCostChangeEvnet(string projectId)
        {
            List<CostChangeEvent> lstForms = new List<CostChangeEvent>();
            //2.取得異動單彙整資料
            using (var context = new topmepEntities())
            {
                //僅針對追加減部分列入 TRANSFLAG='1'
                logger.Debug("query by project and remark:" + projectId);
                string sql = @"SELECT FORM_ID,REMARK,SETTLEMENT_DATE,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM PLAN_COSTCHANGE_ITEM WHERE FORM_ID = f.FORM_ID ) TotalAmt,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM PLAN_COSTCHANGE_ITEM WHERE FORM_ID = f.FORM_ID AND TRANSFLAG='1') RecognizeAmt,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM[PLAN_COSTCHANGE_ITEM] WHERE FORM_ID = f.FORM_ID AND ITEM_QUANTITY> 0 AND TRANSFLAG='1') AddAmt,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM[PLAN_COSTCHANGE_ITEM] WHERE FORM_ID = f.FORM_ID AND ITEM_QUANTITY< 0 AND TRANSFLAG='1') CutAmt
                            FROM PLAN_COSTCHANGE_FORM f WHERE PROJECT_ID=@projectId AND STATUS IN ('進行採購'); ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", projectId));
                logger.Debug("SQL:" + sql);
                lstForms = context.Database.SqlQuery<CostChangeEvent>(sql, parameters.ToArray()).ToList();
            }
            return lstForms;
        }
    }
    //成本異動Service Layer
    public class CostChangeService : PlanService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string strSerialNoKey = "CC";
        public PLAN_COSTCHANGE_FORM form = null;
        public List<PLAN_COSTCHANGE_ITEM> lstItem = null;
        //建立異動單
        public string createChangeOrder(PLAN_COSTCHANGE_FORM form, List<PLAN_COSTCHANGE_ITEM> lstItem)
        {
            //1.新增成本異動單
            SerialKeyService skService = new SerialKeyService();
            form.FORM_ID = skService.getSerialKey(strSerialNoKey);
            form.STATUS = "新建立";
            PLAN_ITEM pi = null;
            int i = 0;
            //2.將資料寫入 
            using (var context = new topmepEntities())
            {
                context.PLAN_COSTCHANGE_FORM.Add(form);
                logger.Debug("create COSTCHANGE_FORM:" + form.FORM_ID);
                foreach (PLAN_COSTCHANGE_ITEM item in lstItem)
                {
                    if (null == item.PLAN_ITEM_ID && "" == item.PLAN_ITEM_ID)
                    {
                        logger.Debug("Object in contract :" + item.PLAN_ITEM_ID);
                        pi = context.PLAN_ITEM.SqlQuery("SELECT * FROM PLAN_ITEM WHERE PLAN_ITEM_ID=@itemId", new SqlParameter("itemId", item.PLAN_ITEM_ID)).First();
                        //補足標單品項欄位
                        if (pi != null && item.ITEM_ID == null)
                        {
                            item.ITEM_ID = pi.ITEM_ID;
                        }
                        if (pi != null && item.ITEM_DESC == null)
                        {
                            item.ITEM_DESC = pi.ITEM_DESC;
                        }
                        if (pi != null && item.ITEM_UNIT == null)
                        {
                            item.ITEM_UNIT = pi.ITEM_UNIT;
                        }
                        if (pi != null && item.ITEM_UNIT_PRICE == null)
                        {
                            item.ITEM_UNIT_PRICE = pi.ITEM_UNIT_PRICE;
                        }
                    }
                    item.FORM_ID = form.FORM_ID;
                    item.PROJECT_ID = form.PROJECT_ID;
                    context.PLAN_COSTCHANGE_ITEM.Add(item);
                    item.CREATE_USER_ID = form.CREATE_USER_ID;
                    item.CREATE_DATE = form.CREATE_DATE;
                    logger.Debug("create COSTCHANGE_ITEM:" + item.PLAN_ITEM_ID);
                }
                i = context.SaveChanges();
            }
            logger.Info("add budget count =" + i);
            return form.FORM_ID;
        }
        //查詢異動單
        public List<PLAN_COSTCHANGE_FORM> getChangeOrders(string projectId, string remark, string status)
        {
            List<PLAN_COSTCHANGE_FORM> lstForms = new List<PLAN_COSTCHANGE_FORM>();
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                logger.Debug("query by project and remark:" + projectId + "," + remark);
                string sql = "SELECT * FROM PLAN_COSTCHANGE_FORM WHERE PROJECT_ID=@projectId";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", projectId));
                if (null != remark && remark != "")
                {
                    sql = sql + " AND REMARK Like @remark";
                    parameters.Add(new SqlParameter("remark", "%" + remark + "%"));
                }
                if (null != status && status != "")
                {
                    sql = sql + " AND STATUS =@status";
                    parameters.Add(new SqlParameter("status", status));
                }
                logger.Debug("SQL:" + sql);
                lstForms = context.PLAN_COSTCHANGE_FORM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            return lstForms;
        }
        //取得單一異動單資料
        public PLAN_COSTCHANGE_FORM getChangeOrderForm(string formId)
        {
            //2.取得異動單資料
            using (var context = new topmepEntities())
            {
                logger.Debug("change form Id:" + formId);
                string sql = "SELECT * FROM PLAN_COSTCHANGE_FORM WHERE FORM_ID=@formId";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", formId));
                logger.Debug("SQL:" + sql);
                form = context.PLAN_COSTCHANGE_FORM.SqlQuery(sql, parameters.ToArray()).First();
                lstItem = form.PLAN_COSTCHANGE_ITEM.ToList();
            }
            project = getProject(form.PROJECT_ID);
            return form;
        }
        //更新異動單資料
        public string updateChangeOrder(PLAN_COSTCHANGE_FORM form, List<PLAN_COSTCHANGE_ITEM> lstItem)
        {
            int i = 0;
            string sqlForm = "UPDATE PLAN_COSTCHANGE_FORM SET REMARK=@remark,SETTLEMENT_DATE=@settlementDate,STATUS=@status,MODIFY_USER_ID=@userId,MODIFY_DATE=@modifyDate WHERE FORM_ID=@formId;";
            string sqlItem = @"UPDATE PLAN_COSTCHANGE_ITEM SET ITEM_DESC=@itemdesc,ITEM_UNIT=@unit,ITEM_UNIT_PRICE=@unitPrice,
                              ITEM_QUANTITY=@Qty,ITEM_REMARK=@remark,TRANSFLAG=@transFlag,MODIFY_USER_ID=@userId,MODIFY_DATE=@modifyDate WHERE ITEM_UID=@uid";
            //2.將資料寫入 

            using (var context = new topmepEntities())
            {
                try
                {
                    //更新表頭
                    context.Database.BeginTransaction();
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("remark", form.REMARK));
                    if (null != form.SETTLEMENT_DATE)
                    {
                        parameters.Add(new SqlParameter("settlementDate", form.SETTLEMENT_DATE));
                    }
                    else
                    {
                        parameters.Add(new SqlParameter("settlementDate", DBNull.Value));
                    }
                    parameters.Add(new SqlParameter("status", form.STATUS));
                    parameters.Add(new SqlParameter("userId", form.MODIFY_USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", form.MODIFY_DATE));
                    parameters.Add(new SqlParameter("formId", form.FORM_ID));
                    i = context.Database.ExecuteSqlCommand(sqlForm, parameters.ToArray());
                    logger.Debug("create COSTCHANGE_FORM:" + sqlForm);
                    foreach (PLAN_COSTCHANGE_ITEM item in lstItem)
                    {
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("itemdesc", item.ITEM_DESC));
                        parameters.Add(new SqlParameter("unit", item.ITEM_UNIT));
                        if (item.ITEM_UNIT_PRICE != null)
                        {
                            parameters.Add(new SqlParameter("unitPrice", item.ITEM_UNIT_PRICE));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("unitPrice", DBNull.Value));
                        }
                        if (item.ITEM_QUANTITY == null)
                        {
                            parameters.Add(new SqlParameter("Qty", DBNull.Value));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("Qty", item.ITEM_QUANTITY));
                        }
                        parameters.Add(new SqlParameter("transFlag", item.TRANSFLAG));
                        parameters.Add(new SqlParameter("remark", item.ITEM_REMARK));
                        parameters.Add(new SqlParameter("userId", item.MODIFY_USER_ID));
                        parameters.Add(new SqlParameter("modifyDate", item.MODIFY_DATE));
                        parameters.Add(new SqlParameter("uid", item.ITEM_UID));
                        i = i + context.Database.ExecuteSqlCommand(sqlItem, parameters.ToArray());
                    }
                    context.Database.CurrentTransaction.Commit();
                }
                catch (Exception ex)
                {
                    context.Database.CurrentTransaction.Rollback();
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                    return "資料更新失敗!!(" + ex.Message + ")";
                }
            }

            return "資料更新成功(" + i + ")!";
        }
        //新增異動單品項
        public int addChangeOrderItem(PLAN_COSTCHANGE_ITEM item)
        {
            int i = 0;
            //2.將資料寫入 
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Debug("create COSTCHANGE_FORM:" + item.FORM_ID);
                    context.PLAN_COSTCHANGE_ITEM.Add(item);
                    i = context.SaveChanges();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
            return i;
        }
        //移除異動單品項
        public int delChangeOrderItem(long itemid)
        {
            int i = 0;
            //2.將品項資料刪除
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = "DELETE FROM PLAN_COSTCHANGE_ITEM WHERE ITEM_UID=@itemUid;";
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("itemUid", itemid));
                    logger.Debug("Delete COSTCHANGE_ITEM:" + itemid);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                    i = context.SaveChanges();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
            return i;
        }
        public string updateChangeOrderStatus(PLAN_COSTCHANGE_FORM form)
        {
            int i = 0;
            string sqlForm = "UPDATE PLAN_COSTCHANGE_FORM SET REMARK=REMARK + @remark,STATUS=@status,MODIFY_USER_ID=@userId,MODIFY_DATE=@modifyDate WHERE FORM_ID=@formId;";
            //2.將資料寫入 
            using (var context = new topmepEntities())
            {
                try
                {
                    //更新表頭
                    context.Database.BeginTransaction();
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("remark", form.REMARK));
                    parameters.Add(new SqlParameter("status", form.STATUS));
                    parameters.Add(new SqlParameter("userId", form.MODIFY_USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", form.MODIFY_DATE));
                    parameters.Add(new SqlParameter("formId", form.FORM_ID));
                    i = context.Database.ExecuteSqlCommand(sqlForm, parameters.ToArray());
                    logger.Debug("create COSTCHANGE_FORM:" + sqlForm);
                    context.Database.CurrentTransaction.Commit();
                }
                catch (Exception ex)
                {
                    context.Database.CurrentTransaction.Rollback();
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                    return "資料更新失敗!!(" + ex.Message + ")";
                }
            }

            return "資料更新成功!(" + i + ")";
        }

    }
}
