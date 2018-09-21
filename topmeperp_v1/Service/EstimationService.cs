using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    /// <summary>
    /// 估驗計價使用
    /// </summary>
    public class EstimationService : PurchaseFormService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string sno_key = "EST";
        //定義驗收單與合約
        string sql4Est = @"with 
esOrder as --驗收單資料
( 
   select Re.PROJECT_ID,Re.PR_ID,Ri.PLAN_ITEM_ID,ri.RECEIPT_QTY,ri.NEED_QTY,ri.ORDER_QTY,ri.REMARK
    from PLAN_PURCHASE_REQUISITION Re
   inner join PLAN_PURCHASE_REQUISITION_ITEM Ri
   on Re.PR_ID=Ri.PR_ID
   and Re.PR_ID Like 'RP%'
   and Re.PROJECT_ID=@projectId
),
contract as (--合約資料
   select f.FORM_NAME,f.SUPPLIER_ID,f.INQUIRY_FORM_ID,ISNULL(f.ISWAGE,'N') TYPE,
   it.PLAN_ITEM_ID,
   it.ITEM_ID,ITEM_DESC,ITEM_QTY,it.ITEM_QTY_ORG,it.ITEM_UNIT,
   it.ITEM_UNIT_PRICE,
   it.ITEM_UNITPRICE_ORG,
   f.PROJECT_ID,
   f.status
   from PLAN_SUP_INQUIRY f
   inner join PLAN_SUP_INQUIRY_ITEM it
   on f.INQUIRY_FORM_ID=it.INQUIRY_FORM_ID
   inner join PLAN_ITEM2_SUP_INQUIRY c
   on c.INQUIRY_FORM_ID=f.INQUIRY_FORM_ID 
   and c.PROJECT_ID=f.PROJECT_ID
   and ISNULL(f.STATUS,'有效')='有效'
   and f.PROJECT_ID=@projectId
)
";
        public ContractModels getContrat(string projectId, string contractid, string prid_s, string prid_e)
        {
            ContractModels c = new ContractModels();
            //1.取得專案資料 
            c.project = getProjectById(projectId);
            //2.取得合約資料
            c.EstimationItems = getEstimationOrder4Expense(projectId, contractid, prid_s, prid_e);
            //3.取得供應商資料
            c.supplier = getSupplierInfo(c.EstimationItems.First().SUPPLIER_ID);
            //4.取得代扣資料
            c.Hold4DeductForm = getPaymentTransfer(projectId,c.supplier.SUPPLIER_ID);
            return c;
        }
        //取得驗收單與相關合約資料
        public List<plansummary> getAllPlanContract(string projectid)
        {
            StringBuilder sb = new StringBuilder(sql4Est);
            List<plansummary> lst = new List<plansummary>();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //parameters.Add(new SqlParameter("supplier", DBNull.Value));
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = @"
select 
ROW_NUMBER() OVER(ORDER BY c.SUPPLIER_ID) AS NO,
c.INQUIRY_FORM_ID,
c.FORM_NAME,c.SUPPLIER_ID,c.TYPE,count(*)　ITEM_ROWS,
MAX(esOrder.PR_ID) as PR_ID_E,MIN(esOrder.PR_ID) as PR_ID_S
  from contract c
inner join esOrder
on c.PLAN_ITEM_ID=esOrder.PLAN_ITEM_ID
group by c.FORM_NAME,c.SUPPLIER_ID,c.TYPE,c.INQUIRY_FORM_ID
                ";
            sql = sb.Append(sql).ToString();
            using (var context = new topmepEntities())
            {
                logger.Debug("get contract sql=" + sql);
                lst = context.Database.SqlQuery<plansummary>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get contract count=" + lst.Count);
            return lst;
        }
        /// <summary>
        /// 取得驗收單明細彙整供估驗單建立相關資料使用
        /// </summary>
        /// <param name="projectid"></param>
        /// <param name="contractId"></param>
        /// <param name="prid_s"></param>
        /// <param name="prid_e"></param>
        /// <returns></returns>
        public List<EstimationItem> getEstimationOrder4Expense(string projectid, string contractId, string prid_s, string prid_e)
        {
            StringBuilder sb = new StringBuilder(sql4Est);
            List<EstimationItem> lst = new List<EstimationItem>();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectId", projectid));
            parameters.Add(new SqlParameter("contractId", contractId));
            parameters.Add(new SqlParameter("prid_s", prid_s));
            parameters.Add(new SqlParameter("prid_e", prid_e));
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = @"
select c.ITEM_ID,c.PLAN_ITEM_ID,
c.ITEM_DESC,
c.ITEM_UNIT,
c.ITEM_QTY AS ITEM_QUANTITY,
c.ITEM_UNIT_PRICE,
SUM(esOrder.RECEIPT_QTY) as EstimationQty,
SUM(RECEIPT_QTY) *c.ITEM_UNIT_PRICE as EstimationAmount,
esOrder.REMARK,
--esOrder.PR_ID,
c.SUPPLIER_ID
  from contract c
inner join esOrder
on c.PLAN_ITEM_ID=esOrder.PLAN_ITEM_ID
where c.PROJECT_ID=@projectId
and c.INQUIRY_FORM_ID=@contractId
and esOrder.PR_ID BETWEEN @prid_s AND @prid_e
GROUP BY c.ITEM_ID,c.PLAN_ITEM_ID,
c.ITEM_DESC,c.ITEM_UNIT,c.ITEM_QTY,c.ITEM_UNIT_PRICE,esOrder.REMARK,c.SUPPLIER_ID
                ";
            sql = sb.Append(sql).ToString();
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql);
                lst = context.Database.SqlQuery<EstimationItem>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        /// <summary>
        /// 建立估驗單與對應的明細資料
        /// </summary>
        /// <param name="form"></param>
        /// <param name="prid_s"></param>
        /// <param name="prid_e"></param>
        public void createEstimationOrder(PLAN_ESTIMATION_FORM form, string prid_s, string prid_e)
        {
            SerialKeyService snoservice = new SerialKeyService();
            form.EST_FORM_ID = snoservice.getSerialKey(sno_key);
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectId", form.PROJECT_ID));
            parameters.Add(new SqlParameter("contractId", form.CONTRACT_ID));
            //parameters.Add(new SqlParameter("EST_FORM_ID", form.EST_FORM_ID));
            parameters.Add(new SqlParameter("prid_s", prid_s));
            parameters.Add(new SqlParameter("prid_e", prid_e));

            StringBuilder sb = new StringBuilder(sql4Est);
            using (var context = new topmepEntities())
            {
                //1,建立主檔
                context.PLAN_ESTIMATION_FORM.Add(form);
                //2.建立驗收單關聯

                string sql4ReceiveOrder = @"
INSERT INTO PLAN_ESTIMATION2PURCHASE
select DISTINCT esOrder.PR_ID AS PR_ID,@EST_FORM_ID AS EXT_FORM_ID
  from contract c
inner join esOrder
on c.PLAN_ITEM_ID=esOrder.PLAN_ITEM_ID
where c.PROJECT_ID=@projectId
and c.INQUIRY_FORM_ID=@contractId
and esOrder.PR_ID BETWEEN @prid_s AND @prid_e 
";
                string sql = sb.Append(sql4ReceiveOrder).Replace("@EST_FORM_ID", "'" + form.EST_FORM_ID + "'").ToString();
                logger.Debug(sql);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());

                //3.建立明細
                string sql4Detail = @"
INSERT INTO PLAN_ESTIMATION_ITEM
select 
@EST_FORM_ID AS EXT_FORM_ID,
c.PLAN_ITEM_ID,
SUM(esOrder.RECEIPT_QTY) as EST_QTY,
1 as EST_RATIO,
SUM(RECEIPT_QTY) *c.ITEM_UNIT_PRICE as EST_AMOUNT,
NULL AS REMARK
 from contract c
inner join esOrder
on c.PLAN_ITEM_ID=esOrder.PLAN_ITEM_ID
where c.PROJECT_ID=@projectId
and c.INQUIRY_FORM_ID=@contractId
and esOrder.PR_ID BETWEEN @prid_s AND @prid_e 
GROUP BY c.PLAN_ITEM_ID,c.ITEM_UNIT_PRICE;
";
                sb = new StringBuilder(sql4Est);
                sql = sb.Append(sql4Detail).Replace("@EST_FORM_ID", "'" + form.EST_FORM_ID + "'").ToString();
                logger.Debug(sql);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                //4.建立扣款明細 todo
                context.SaveChanges();
            }
        }
        //取得代付扣款彙整資料
        private List<Model4PaymentTransfer> getPaymentTransfer(string projectId, string supplierId)
        {
            List<Model4PaymentTransfer> lstSummary = null;
            string sql = @"
select f.EST_FORM_ID,
 f.PROJECT_ID,
 f.CONTRACT_ID,
 SUM(it.[EST_AMOUNT]) PAID_AMOUNT,
 f.HOLD4REMARK
from [PLAN_ESTIMATION_FORM] f
inner  join PLAN_ESTIMATION_ITEM it
on f.EST_FORM_ID=it.EST_FORM_ID
where f.PROJECT_ID=@projectId
and f.hold4supplier=@supplierId
group by f.EST_FORM_ID,f.PROJECT_ID,f.CONTRACT_ID,f.HOLD4REMARK
";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectId", projectId));
            parameters.Add(new SqlParameter("supplierId", supplierId));

            using (var context = new topmepEntities())
            {
                lstSummary = context.Database.SqlQuery<Model4PaymentTransfer>(sql, parameters.ToArray()).ToList();
            }
            return lstSummary;
        }
        public List<EstimationForm> getContractItemById(string contractid, string projectid)
        {

            logger.Info("get contract item by contractid  =" + contractid);
            List<EstimationForm> lstItem = new List<EstimationForm>();
            //處理SQL 預先填入合約代號,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<EstimationForm>("SELECT pi.*, psi.ITEM_QTY AS mapQty, A.CUM_QTY AS CUM_EST_QTY, ISNULL(B.CUM_QTY, 0) AS CUM_RECPT_QTY, ISNULL(B.CUM_QTY, 0)-ISNULL(A.CUM_QTY,0) AS Quota FROM PLAN_ITEM pi " +
                    "LEFT JOIN (SELECT PLAN_ITEM_ID, ITEM_QTY FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID =@contractid)psi ON pi.PLAN_ITEM_ID = psi.PLAN_ITEM_ID LEFT JOIN (SELECT ei.PLAN_ITEM_ID, SUM(ei.EST_QTY) AS CUM_QTY " +
                    "FROM PLAN_ESTIMATION_ITEM ei LEFT JOIN PLAN_ESTIMATION_FORM ef ON ei.EST_FORM_ID = ef.EST_FORM_ID " +
                    "WHERE ef.CONTRACT_ID = @contractid GROUP BY ei.PLAN_ITEM_ID)A ON pi.PLAN_ITEM_ID = A.PLAN_ITEM_ID " +
                    "LEFT JOIN (SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS CUM_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr " +
                    "ON pri.PR_ID = pr.PR_ID WHERE pri.PR_ID LIKE 'RP%' AND pr.PROJECT_ID = @projectid GROUP BY pri.PLAN_ITEM_ID)B ON pi.PLAN_ITEM_ID = B.PLAN_ITEM_ID WHERE " +
                    "pi.INQUIRY_FORM_ID = @contractid OR pi.MAN_FORM_ID = @contractid ; "
            , new SqlParameter("contractid", contractid), new SqlParameter("projectid", projectid)).ToList();
            }

            return lstItem;
        }

        //取得個別材料廠商合約資料與金額
        public plansummary getPlanContract4Est(string contractid)
        {
            plansummary lst = new plansummary();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT A.INQUIRY_FORM_ID AS CONTRACT_ID, A.SUPPLIER_ID, A.FORM_NAME, " +
                    "SUM(A.formQty * A.ITEM_UNIT_COST) MATERIAL_COST, SUM(A.formQty * ISNULL(A.MAN_PRICE, 0)) WAGE_COST, " +
                    "SUM(A.ITEM_QUANTITY * A.ITEM_UNIT_PRICE) REVENUE, SUM(A.mapQty * A.tndPrice * A.BUDGET_RATIO / 100) BUDGET, " +
                    "(SUM(A.formQty * A.ITEM_UNIT_COST) + SUM(A.formQty * ISNULL(A.MAN_PRICE, 0))) COST, (SUM(A.ITEM_QUANTITY * A.ITEM_UNIT_PRICE) - " +
                    "SUM(A.formQty * A.ITEM_UNIT_COST) - SUM(A.formQty * ISNULL(A.MAN_PRICE, 0))) PROFIT, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY A.SUPPLIER_ID) AS NO FROM (SELECT pi.*, s.SUPPLIER_ID AS ID, psi.ITEM_QTY AS formQty, map.QTY AS mapQty, tpi.ITEM_UNIT_PRICE AS tndPrice FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                    "pi.SUPPLIER_ID = s.COMPANY_NAME LEFT JOIN (SELECT PLAN_ITEM_ID, ITEM_QTY FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID =@contractid)psi ON pi.PLAN_ITEM_ID = psi.PLAN_ITEM_ID " +
                    "LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT_ITEM tpi ON pi.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID)A GROUP BY A.PROJECT_ID, A.INQUIRY_FORM_ID, A.FORM_NAME, A.SUPPLIER_ID HAVING A.INQUIRY_FORM_ID =@contractid ; "
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
                lst = context.Database.SqlQuery<plansummary>("SELECT  A.MAN_FORM_ID AS CONTRACT_ID, A.MAN_SUPPLIER_ID, A.MAN_FORM_NAME, " +
                    "SUM(A.formQty * ISNULL(A.MAN_PRICE, 0)) WAGE_COST, " +
                    "count(*) AS ITEM_ROWS, ROW_NUMBER() OVER(ORDER BY A.MAN_SUPPLIER_ID) AS NO FROM(SELECT pi.*, s.SUPPLIER_ID AS ID, psi.ITEM_QTY AS formQty FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                    "pi.MAN_SUPPLIER_ID = s.COMPANY_NAME LEFT JOIN (SELECT PLAN_ITEM_ID, ITEM_QTY FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID =@contractid)psi ON pi.PLAN_ITEM_ID = psi.PLAN_ITEM_ID)A GROUP BY A.PROJECT_ID, A.MAN_SUPPLIER_ID, A.MAN_FORM_NAME, A.MAN_FORM_ID HAVING A.MAN_FORM_ID =@contractid ; "
                   , new SqlParameter("contractid", contractid)).First();
            }
            return lst;
        }
    }
}