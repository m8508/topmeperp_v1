using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
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
        #region 定義驗收單與合約
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
        #endregion
        public ContractModels getContract(string projectId, string contractid, string prid_s, string prid_e)
        {
            ContractModels c = new ContractModels();

            PLAN_ESTIMATION_FORM form = new PLAN_ESTIMATION_FORM();
            c.planEST = form;
            //1.取得專案資料 
            logger.Debug("get project by project_id=" + projectId);
            c.project = getProjectById(projectId);
            //2.取得合約資料
            logger.Debug("get EstimationOrder4Expense by project_id=" + contractid + "," + prid_s + "," + prid_e);
            c.EstimationItems = getEstimationOrder4Expense(projectId, contractid, prid_s, prid_e);
            //3.取得供應商資料
            logger.Debug("get Supplier by SUPPLIER_ID=" + c.EstimationItems.First().SUPPLIER_ID);
            c.supplier = getSupplierInfo(c.EstimationItems.First().SUPPLIER_ID);
            //4.取得代扣資料
            logger.Debug("get Hold4DeductForm by SUPPLIER_ID=" + c.supplier.SUPPLIER_ID);
            c.Hold4DeductForm = getPaymentTransfer(projectId, c.supplier.SUPPLIER_ID);
            return c;
        }
        /// <summary>
        /// 取得估驗單資料
        /// </summary>
        /// <param name="estimationOrderId"></param>
        public ContractModels getEstimationOrder(string formId)
        {
            ContractModels c = new ContractModels();
            //3.取得估驗單資料
            c.planEST = getForm(formId);
            //1.取得專案資料 
            c.project = getProjectById(c.planEST.PROJECT_ID);
            //2.合約資料 
            PurchaseFormService s = new PurchaseFormService();
            s.getInqueryForm(c.planEST.CONTRACT_ID);
            c.supContract = s.formInquiry;
            c.supContractItems = s.formInquiryItem;

            //4.取得估驗明細資料
            c.EstimationItems = getItems(c.planEST.EST_FORM_ID);
            //5.取得供應商資料
            c.supplier = getSupplierInfo(c.planEST.PAYEE);
            //5.1 取得憑證資料
            c.EstimationInvoices = getEstimationInvoice(c.planEST.EST_FORM_ID);
            //6.取得代付支出
            c.EstimationHoldPayments = getHoldPayment(formId);
            //7.取得代扣支出
            c.Hold4DeductForm = getPaymentTransfer(formId);
            //8.取得任務清單
            Flow4EstimationForm wkservice = new Flow4EstimationForm();
            wkservice.getTask(formId);
            c.task = wkservice.task.task;
            return c;
        }
        //取得估驗單清單
        public List<EstimationOrderForm> getFormList(string projectid, string status)
        {
            List<EstimationOrderForm> lst = null;
            string sql = @"
select p.PROJECT_NAME,c.SUPPLIER_ID SUPPLIER_NAME,c.FORM_NAME CONTRACT_NAME,f.*
from PLAN_ESTIMATION_FORM f
inner join TND_PROJECT p on f.project_id=p.project_id
inner join PLAN_SUP_INQUIRY c on f.CONTRACT_ID=c.INQUIRY_FORM_ID
WHERE  f.PROJECT_ID=@projectid AND (@status is null or f.STATUS=@status)
";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            if (null == status)
            {
                parameters.Add(new SqlParameter("status", DBNull.Value));
            }
            else
            {
                parameters.Add(new SqlParameter("status", status));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql);
                lst = context.Database.SqlQuery<EstimationOrderForm>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //取得估驗單主檔
        public PLAN_ESTIMATION_FORM getForm(string formId)
        {
            PLAN_ESTIMATION_FORM form = null;
            using (var context = new topmepEntities())
            {
                form = context.PLAN_ESTIMATION_FORM.Where(f => f.EST_FORM_ID == formId).FirstOrDefault();
            }
            return form;
        }
        //取得供應商查詢條件
        private string getSupplierByEstimationOrder(string formId)
        {
            List<string> supplierId = new List<string>();
            string sql = @"
select Distinct SUPPLIER_ID from PLAN_ESTIMATION2PURCHASE e2p
inner join PLAN_PURCHASE_REQUISITION  pr
on e2p.PR_ID=pr.PR_ID
where e2p.EST_FORM_ID=@formId
";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", formId));
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql);
                supplierId = context.Database.SqlQuery<string>(sql, parameters.ToArray()).ToList();
            }
            if (supplierId.Count > 1)
            {
                logger.Error("Data Error :" + formId);
            }
            return supplierId[0];
        }
        //取得估驗單明細資料
        public List<EstimationItem> getItems(string formId)
        {
            List<EstimationItem> lst = new List<EstimationItem>();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", formId));
            string sql = @"
select 
f.CONTRACT_ID CONTRACT_ID,f.EST_FORM_ID EST_FORM_ID,
it.PLAN_ITEM_ID,cit.ITEM_ID,cit.ITEM_DESC,cit.ITEM_UNIT,cit.ITEM_UNIT_PRICE,
 cit.ITEM_QTY as ITEM_QUANTITY,it.EST_QTY EstimationQty,it.EST_AMOUNT EstimationAmount,
 it.REMARK,it.EST_ITEM_ID
from  PLAN_ESTIMATION_FORM f
inner join PLAN_ESTIMATION_ITEM it
on f.EST_FORM_ID=it.EST_FORM_ID 
inner join PLAN_SUP_INQUIRY_ITEM cit
on it.PLAN_ITEM_ID=cit.PLAN_ITEM_ID 
and f.CONTRACT_ID=cit.INQUIRY_FORM_ID
and f.EST_FORM_ID=@formId
";
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql);
                lst = context.Database.SqlQuery<EstimationItem>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //處理廠商請款表單發票明細資料
        public static List<PLAN_ESTIMATION_INVOICE> getSellInvoice(FormCollection f, string formId, string contracId, List<PLAN_ESTIMATION_INVOICE> lstInvoice, string Invoices)
        {
            string[] aryInvoiceId = Invoices.Split(',');
            string[] aryInvoiceAmount = (string[])f.GetValue("invoiceAmt").RawValue;
            string[] aryInvoiceTax = (string[])f.GetValue("invoiceTax").RawValue;
            string[] aryInvoiceDate = f["invoiceDate"].Split(',');
            string[] aryInvoiceType = f["invoicetype"].Split(',');
            string[] aryInvoiceNote = (string[])f.GetValue("invoiceNote").RawValue;
            if (null != Invoices)
            {
                lstInvoice = new List<PLAN_ESTIMATION_INVOICE>();
                string[] aryInvoiceNo = Invoices.Split(',');
                for (int i = 0; i < aryInvoiceId.Length; i++)
                {
                    PLAN_ESTIMATION_INVOICE invoice = new PLAN_ESTIMATION_INVOICE();
                    invoice.EST_FORM_ID = formId;
                    invoice.CONTRACT_ID = contracId;
                    //發票號碼
                    invoice.INVOICE_NUMBER = aryInvoiceNo[i];
                    //發票金額
                    if (aryInvoiceAmount[i] != "")
                    {
                        invoice.AMOUNT = decimal.Parse(aryInvoiceAmount[i]);
                    }
                    //發票類型
                    invoice.TYPE = aryInvoiceType[i];
                    //發票稅金
                    if (aryInvoiceTax[i] != "")
                    {
                        invoice.TAX = decimal.Parse(aryInvoiceTax[i]);
                    }
                    else
                    {
                        //若使用這位輸入，依據憑證類型計算稅金
                        InvoiceService s = new InvoiceService();
                        invoice.TAX = Convert.ToDecimal(Math.Round(Convert.ToDouble(s.getTaxRate(invoice.TYPE) * invoice.AMOUNT), 0, MidpointRounding.AwayFromZero));
                    }
                    invoice.INVOICE_DATE = DateTime.Parse(aryInvoiceDate[i]);

                    invoice.NOTE = aryInvoiceNote[i];
                    lstInvoice.Add(invoice);
                }
            }
            return lstInvoice;
        }
        //取得估驗請款憑證資料
        public List<PLAN_ESTIMATION_INVOICE> getEstimationInvoice(string formId)
        {
            List<PLAN_ESTIMATION_INVOICE> lst = new List<PLAN_ESTIMATION_INVOICE>();
            using (var context = new topmepEntities())
            {
                lst = context.PLAN_ESTIMATION_INVOICE.Where(x => x.EST_FORM_ID == formId).Select(x => x).ToList();
                logger.Debug("get invoice by formId=" + formId + ",count=" + lst.Count);
            }
            return lst;
        }
        //取得代付支出明細資料
        public List<PLAN_ESTIMATION_HOLDPAYMENT> getHoldPayment(string formId)
        {
            List<PLAN_ESTIMATION_HOLDPAYMENT> lst = new List<PLAN_ESTIMATION_HOLDPAYMENT>();
            logger.Debug("get PLAN_ESTIMATION_HOLDPAYMENT by " + formId);
            using (var context = new topmepEntities())
            {
                lst = context.PLAN_ESTIMATION_HOLDPAYMENT.Where(f => f.EST_FORM_ID == formId).ToList();
            }
            return lst;
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
        public void createEstimationOrder(PLAN_ESTIMATION_FORM form,
            List<PLAN_ESTIMATION_HOLDPAYMENT> lstHoldPayment,
            List<PLAN_ESTIMATION_PAYMENT_TRANSFER> listTransferPayment,
            List<PLAN_ESTIMATION_INVOICE> listInvoice,
            string prid_s, string prid_e)
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
                //3.1 更新表單檔頭金額
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
UPDATE 
PLAN_ESTIMATION_FORM
SET PAID_AMOUNT = 
(SELECT SUM(EST_AMOUNT) FROM PLAN_ESTIMATION_ITEM  WHERE EST_FORM_ID=@EST_FORM_ID)
WHERE EST_FORM_ID=@EST_FORM_ID;
";
                sb = new StringBuilder(sql4Est);
                sql = sb.Append(sql4Detail).Replace("@EST_FORM_ID", "'" + form.EST_FORM_ID + "'").ToString();
                logger.Debug(sql);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                //3.1 付款憑證資料
                modifyEstimationInvoice(form, listInvoice, context);
                //4.建立代付資料
                modifyEstimationHold(form, lstHoldPayment, context);
                logger.Debug("get Hold4Payment=" + JsonConvert.SerializeObject(lstHoldPayment).ToString());
                //5.建立代付扣款明細 
                modifyEstimationTransfer(form, listTransferPayment, context);
                logger.Debug("get TransferPayment=" + JsonConvert.SerializeObject(listTransferPayment).ToString());
                context.SaveChanges();
            }
            //6.建立彙整金額
            SumEstimationForm(form);
        }
        /// <summary>
        /// 將估驗單相關金額，加入表頭內
        /// </summary>
        /// <param name="formId"></param>
        private void SumEstimationForm(PLAN_ESTIMATION_FORM f)
        {
            //TODO : 保留款、預付款、其他扣款尚未計算
            string sql = @"
 UPDATE PLAN_ESTIMATION_FORM
 SET 
 PAID_AMOUNT=INV_ALL.PAID_AMOUNT-INV_ALL.PAYMENT_TRANSFER ,-- 應付金額(須扣除代付支出)
 TAX_AMOUNT=INV_ALL.TAX_AMOUNT,--營業稅
  PAYMENT_TRANSFER=INV_ALL.PAYMENT_TRANSFER,--代付支出
  PAYMENT_DEDUCTION=INV_ALL.PAYMENT_DEDUCTION,----代付扣回
  RETENTION_PAYMENT=0,--保留款(需另外計算)
  OTHER_PAYMENT=0, -- 其他扣款(需另外計算)
  PREPAY_AMOUNT=0, --預付款(需另外計算)
  FOREIGN_PAYMENT=0 --外勞款(尚無資料)
  FROM 
  (SELECT INV.EST_FORM_ID FORM_ID
,ISNULL(SUM(HOLD.HOLD_AMOUNT),0) PAYMENT_TRANSFER  --代付支出
,SUM(INV.AMOUNT) PAID_AMOUNT  -- 應付金額
,SUM(INV.TAX) TAX_AMOUNT  --營業稅
,ISNULL(SUM(TRF.PAID_AMOUNT),0) PAYMENT_DEDUCTION   --代付扣回
FROM PLAN_ESTIMATION_INVOICE INV
LEFT JOIN PLAN_ESTIMATION_HOLDPAYMENT HOLD
ON INV.EST_FORM_ID=HOLD.EST_FORM_ID
LEFT JOIN PLAN_ESTIMATION_PAYMENT_TRANSFER TRF
ON INV.EST_FORM_ID=TRF.TRANSFER_FORM_ID
 WHERE INV.EST_FORM_ID = @formId 
 AND INV.CONTRACT_ID=@contractId
 GROUP BY INV.EST_FORM_ID
 ) INV_ALL
 WHERE INV_ALL.FORM_ID= PLAN_ESTIMATION_FORM.EST_FORM_ID
";
            logger.Debug(sql + ",formId=" + f.EST_FORM_ID +",Contract=" + f.CONTRACT_ID);
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", f.EST_FORM_ID));
            parameters.Add(new SqlParameter("contractId", f.CONTRACT_ID));
            using (var context = new topmepEntities())
            {
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
        }
        /// <summary>
        /// 建立發票資料
        /// </summary>
        public static void modifyEstimationInvoice(PLAN_ESTIMATION_FORM form, List<PLAN_ESTIMATION_INVOICE> lstInvoice, topmepEntities context)
        {
            string sql = "DELETE PLAN_ESTIMATION_INVOICE WHERE EST_FORM_ID=@formId";
            logger.Debug("remove estimation invoice:sql=" + sql + ",formId=" + form.EST_FORM_ID);
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", form.EST_FORM_ID));
            context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            logger.Debug("sql=" + sql + ",formId=" + form.EST_FORM_ID);

            if (null != lstInvoice && lstInvoice.Count > 0)
            {
                foreach (PLAN_ESTIMATION_INVOICE inv in lstInvoice)
                {
                    inv.EST_FORM_ID = form.EST_FORM_ID;
                    context.PLAN_ESTIMATION_INVOICE.Add(inv);
                }
            }
        }
        /// <summary>
        /// 建立代付扣回資料
        /// </summary>
        /// <param name="form"></param>
        /// <param name="listTransferPayment"></param>
        /// <param name="context"></param>
        private static void modifyEstimationTransfer(PLAN_ESTIMATION_FORM form, List<PLAN_ESTIMATION_PAYMENT_TRANSFER> listTransferPayment, topmepEntities context)
        {
            string sql = "DELETE PLAN_ESTIMATION_PAYMENT_TRANSFER WHERE PAYMENT_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", form.EST_FORM_ID));
            context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            logger.Debug("sql=" + sql + ",formId=" + form.EST_FORM_ID);

            if (null != listTransferPayment && listTransferPayment.Count > 0)
            {
                foreach (PLAN_ESTIMATION_PAYMENT_TRANSFER trans in listTransferPayment)
                {
                    trans.PAYMENT_FORM_ID = form.EST_FORM_ID;
                    context.PLAN_ESTIMATION_PAYMENT_TRANSFER.Add(trans);
                }
            }
        }

        /// <summary>
        /// 建立代付款資料
        /// </summary>
        /// <param name="form"></param>
        /// <param name="lstHoldPayment"></param>
        /// <param name="context"></param>
        private static void modifyEstimationHold(PLAN_ESTIMATION_FORM form, List<PLAN_ESTIMATION_HOLDPAYMENT> lstHoldPayment, topmepEntities context)
        {
            //1.delete exist data
            string sql = "DELETE PLAN_ESTIMATION_HOLDPAYMENT WHERE EST_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", form.EST_FORM_ID));
            context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            logger.Debug("sql=" + sql + ",formId=" + form.EST_FORM_ID);

            //2.建立代付款資料
            if (lstHoldPayment != null && lstHoldPayment.Count > 0)
            {
                foreach (PLAN_ESTIMATION_HOLDPAYMENT hold in lstHoldPayment)
                {
                    hold.EST_FORM_ID = form.EST_FORM_ID;
                    context.PLAN_ESTIMATION_HOLDPAYMENT.Add(hold);
                }
            }
        }
        /// <summary>
        /// 修改估驗單資料
        /// </summary>
        /// <param name="form"></param>
        /// <param name="lstHoldPayment"></param>
        /// <param name="listTransferPayment"></param>
        public void modifyEstimationOrder(PLAN_ESTIMATION_FORM form,
            List<PLAN_ESTIMATION_HOLDPAYMENT> lstHoldPayment,
            List<PLAN_ESTIMATION_PAYMENT_TRANSFER> listTransferPayment,
            List<PLAN_ESTIMATION_INVOICE> listInvoice)
        {
            using (var context = new topmepEntities())
            {
                modifyEstimationForm(context, form);
                modifyEstimationInvoice(form, listInvoice, context);
                modifyEstimationHold(form, lstHoldPayment, context);
                modifyEstimationTransfer(form, listTransferPayment, context);
                context.SaveChanges();
            }
            //.建立彙整金額
            SumEstimationForm(form);
        }
        /// <summary>
        /// 刪除估驗單資料
        /// </summary>
        /// <param name="formId"></param>
        public void delEstimationOrder(string formId)
        {
            StringBuilder sb = new StringBuilder();
            //刪除 估驗單明細 PLAN_ESTIMATION_ITEM 
            sb.Append("DELETE PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID=@formId;");
            //刪除 代付款紀錄 PLAN_ESTIMATION_HOLDPAYMENT
            sb.Append("DELETE PLAN_ESTIMATION_HOLDPAYMENT WHERE EST_FORM_ID=@formId;");
            //刪除 代付扣款紀錄 PLAN_ESTIMATION_PAYMENT_TRANSFER
            sb.Append("DELETE PLAN_ESTIMATION_PAYMENT_TRANSFER WHERE PAYMENT_FORM_ID=@formId;");
            //刪除 估驗單資料 PLAN_ESTIMATION_FORM
            sb.Append("DELETE PLAN_ESTIMATION_FORM WHERE EST_FORM_ID=@formId;");
            //刪除 驗收單與估驗單對照資料 PLAN_ESTIMATION2PURCHASE
            sb.Append("DELETE PLAN_ESTIMATION2PURCHASE WHERE EST_FORM_ID=@formId;");

            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", formId));
                string sql = sb.ToString();
                logger.Debug("sql=" + sql + ",formId=" + formId);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
        }
        //修改基本資料
        private static void modifyEstimationForm(topmepEntities context, PLAN_ESTIMATION_FORM form)
        {
            string sql = @"
UPDATE PLAN_ESTIMATION_FORM
   SET PROJECT_ID = @projectid
      ,CONTRACT_ID = @contractid
      ,PLUS_TAX = @plustax
      ,TAX_AMOUNT = @taxamount
      ,PAYMENT_TRANSFER = @paymenttransfer
      ,PAID_AMOUNT = @paidamount
      ,FOREIGN_PAYMENT = @foreignpayment
      ,RETENTION_PAYMENT = @retentionpayment
      ,REMARK = @remark
      ,SETTLEMENT = @settlement
      ,TYPE = @type
      ,STATUS = @status
      ,TAX_RATIO = @taxration
      ,MODIFY_DATE = getdate()
      ,INVOICE = @invoice
      ,REJECT_DESC = @rejectdesc
      ,PROJECT_NAME = @projectname
      ,PAYEE = @payee
      ,PAYMENT_DATE = @paymentDate
      ,INDIRECT_COST_TYPE = @indirectCostType
 WHERE EST_FORM_ID = @formId
                ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", form.EST_FORM_ID));
            parameters.Add(new SqlParameter("projectid", form.PROJECT_ID));
            parameters.Add(new SqlParameter("contractid", form.CONTRACT_ID));
            parameters.Add(new SqlParameter("plustax", form.PLUS_TAX ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("taxamount", form.TAX_AMOUNT ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("paymenttransfer", form.PAYMENT_TRANSFER ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("paidamount", form.PAID_AMOUNT ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("foreignpayment", form.FOREIGN_PAYMENT ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("retentionpayment", form.RETENTION_PAYMENT ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("remark", form.REMARK ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("settlement", form.SETTLEMENT ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("type", form.TYPE ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("status", form.STATUS ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("taxration", form.TAX_RATIO ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("invoice", form.INVOICE ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("rejectdesc", form.REJECT_DESC ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("projectname", form.PROJECT_NAME ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("payee", form.PAYEE ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("paymentDate", form.PAYMENT_DATE ?? (object)DBNull.Value));
            parameters.Add(new SqlParameter("indirectCostType", form.INDIRECT_COST_TYPE ?? (object)DBNull.Value));
            context.Database.ExecuteSqlCommand(sql, parameters.ToArray());

            //   parameters.Add(new SqlParameter("supplierId", supplierId));

        }

        //取得代付扣款彙整資料--依據驗收單相關取得對應的代付資料
        private List<Model4PaymentTransfer> getPaymentTransfer(string projectId, string supplierId)
        {
            List<Model4PaymentTransfer> lstSummary = null;
            string sql = @"
SELECT 
*
 FROM
(
--代付資料
SELECT F.PROJECT_ID,F.CONTRACT_ID,F.PAYEE,H.* FROM 
PLAN_ESTIMATION_FORM F INNER JOIN
PLAN_ESTIMATION_HOLDPAYMENT H
ON F.EST_FORM_ID=H.EST_FORM_ID
AND F.PROJECT_ID=@projectId
) A
WHERE LEFT(SUPPLIER_ID,7) =@supplierId
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
        //取得代付扣款資料
        private List<Model4PaymentTransfer> getPaymentTransfer(string formId)
        {
            List<Model4PaymentTransfer> lstSummary = null;
            string sql = @"
select 
PAYMENT_FORM_ID EST_FORM_ID,f.PROJECT_ID,S.COMPANY_NAME,f.CONTRACT_ID,f.PAYEE,f.CREATE_DATE,
tf.* from　PLAN_ESTIMATION_PAYMENT_TRANSFER tf
INNER JOIN PLAN_ESTIMATION_FORM  f
ON tf.PAYMENT_FORM_ID=f.EST_FORM_ID
INNER JOIN TND_SUPPLIER S
ON f.PAYEE=s.SUPPLIER_ID
WHERE tf.PAYMENT_FORM_ID=@formId
";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", formId));
            logger.Debug("sql=" + sql + ",formId=" + formId);
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
    //發票(憑證)邏輯作業服務
    public class InvoiceService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //憑證類型陣列
        string[] InvoiceType = { "二聯式", "三聯式", "收據", "工資單", "對開發票", "折讓單-溢開折讓", "折讓單-扣款折讓", "其他扣款" };
        //對應稅率資料
        decimal[] TaxRate = { 0.0M, 0.05M, 0.0M, 0.0M, 0.0M, 0.0M, 0.0M, 0.0M };
        public InvoiceService()
        {

        }
        public string[] getInvoiceType()
        {
            return InvoiceType;
        }
        public decimal getTaxRate(string invoicetype)
        {
            int idx = Array.IndexOf(InvoiceType, invoicetype);
            if (idx > -1)
            {
                logger.Debug("get tax Rate Type=" + invoicetype + ",Rate=" + TaxRate[idx]);
                return TaxRate[idx];
            }
            return -1;
        }
    }
}