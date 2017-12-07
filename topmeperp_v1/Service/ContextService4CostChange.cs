using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
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
        public List<PLAN_COSTCHANGE_FORM> getChangeOrders(string projectId, string remark)
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
                    sql = sql + " AND REMARK Like @remark;";
                    parameters.Add(new SqlParameter("remark", "'%" + remark + "%'"));
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
            string sqlForm = "UPDATE PLAN_COSTCHANGE_FORM SET REMARK=@remark,STATUS=@status,MODIFY_USER_ID=@userId,MODIFY_DATE=@modifyDate WHERE FORM_ID=@formId;";
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

            return "資料更新成功!(" + i + ")";
        }
    }
}