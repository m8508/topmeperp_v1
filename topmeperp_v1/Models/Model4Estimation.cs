using System;
using System.Collections.Generic;

namespace topmeperp.Models
{
    /// <summary>
    /// 估驗單管理與審核表單管理使用
    /// </summary>
    public class EstimationFormApprove
    {
        public List<ExpenseFlowTask> lstEstimationFlowTask { set; get; }
        public List<EstimationOrderForm> lstEstimationForm { set; get; }
    }
    public class purchasesummary
    {
        public string FORM_NAME { get; set; }
        public string INQUIRY_FORM_ID { get; set; }
        public string SUPPLIER_ID { get; set; }
        public Nullable<int> TOTALROWS { get; set; }
        public Nullable<decimal> TAmount { get; set; }
        public Nullable<decimal> Budget { get; set; }
        public string STATUS { get; set; }
    }
    #region 估驗、請款物件
    /// <summary>
    /// 廠商合約物件
    /// </summary>
    public class ContractModels
    {
        //專案資料
        public TND_PROJECT project { get; set; }
        //供應商(廠商資料)
        public TND_SUPPLIER supplier { get; set; }
        //估驗單表頭資料
        public PLAN_ESTIMATION_FORM planEST { get; set; }
        //估驗單表身資料
        public IEnumerable<EstimationForm> planESTItem { get; set; }
        //合約(估驗)項目清單
        public IEnumerable<EstimationItem> EstimationItems { get; set; }
        //發票資料
        public IEnumerable<PLAN_ESTIMATION_INVOICE> EstimationInvoices { get; set; }
        //代付支出明細資料
        public IEnumerable<PLAN_ESTIMATION_HOLDPAYMENT> EstimationHoldPayments { get; set; }
        //代付扣款明細
        public IEnumerable<Model4PaymentTransfer> Hold4DeductForm { get; set; }
        public IEnumerable<plansummary> contractItems { get; set; }
        public IEnumerable<plansummary> wagecontractItems { get; set; }
        //public IEnumerable<PLAN_PAYMENT_TERMS> paymentTerms { get; set; }
       // public PaymentTermsFunction planpayment { get; set; }
       // public IEnumerable<PLAN_ITEM> planItems { get; set; }

        public IEnumerable<PURCHASE_ORDER> planOrder { get; set; }
        public IEnumerable<RevenueFromOwner> ownerConFile { get; set; }
        //工作流程定義物件
        public ExpenseFlowTask task { get; set; }
    }
    public class EstimationOrderForm: PLAN_ESTIMATION_FORM
    {
        public string FORM_NAME { get; set; }
        public string PROJECT_NAME { get; set; }
        public string CONTRACT_NAME { get; set; }
        public string SUPPLIER_NAME { get; set; }
        //對應合約特定期間的驗收單(起始單號)
        public string PR_ID_S { get; set; }
        //對應合約特定期間的驗收單(結束單號)
        public string PR_ID_E { get; set; }
    }
    //估驗明細資料
    public class EstimationItem : PLAN_ITEM
    {
        public string SUPPLIER_ID { get; set; }
        //前期數量
        public Nullable<decimal> PriorQty { get; set; }
        //前期金額
        public Nullable<decimal> PriorAmount { get; set; }

        //本期數量
        public Nullable<decimal> EstimationQty { get; set; }
        //本期金額
        public Nullable<decimal> EstimationAmount { get; set; }
        //累計數量
        public Nullable<decimal> TotalQty { get; set; }
        //累計金額
        public Nullable<decimal> TotalAmount { get; set; }
        public Nullable<long> EST_ITEM_ID { get; set; }
    }
    /// <summary>
    /// 估驗請款彙總紀錄
    /// </summary>
    public class plansummary
    {
        public string FORM_NAME { get; set; }
        public Nullable<int> ITEM_ROWS { get; set; }
        public string SUPPLIER_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public Nullable<decimal> REVENUE { get; set; }
        public Nullable<decimal> COST { get; set; }
        public Nullable<decimal> BUDGET { get; set; }
        public Nullable<decimal> PROFIT { get; set; }
        public Int64 NO { get; set; }
        public Nullable<decimal> TOTAL_REVENUE { get; set; }
        public Nullable<decimal> TOTAL_BUDGET { get; set; }
        public Nullable<decimal> TOTAL_COST { get; set; }
        public Nullable<decimal> TOTAL_PROFIT { get; set; }
        public Nullable<decimal> MATERIAL_COST { get; set; }
        public Nullable<decimal> WAGE_COST { get; set; }
        public string MAN_FORM_NAME { get; set; }
        public string MAN_SUPPLIER_ID { get; set; }
        public string CONTRACT_NAME { get; set; }
        public string TYPE { get; set; }
        public Nullable<decimal> WAGE_BUDGET { get; set; }
        public string INQUIRY_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }

        //對應合約特定期間的驗收單(起始單號)
        public string PR_ID_S { get; set; }
        //對應合約特定期間的驗收單(結束單號)
        public string PR_ID_E { get; set; }
    }
    #endregion 估驗、請款物件
    //代付扣回資料
    public class Model4PaymentTransfer:PLAN_ESTIMATION_PAYMENT_TRANSFER
    {
        public string EST_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public DateTime CREATE_DATE { get; set; }
        public string COMPANY_NAME { get; set; }
        public string PAYEE { get; set; }
    }
    public class ESTFunction
    {
        public Int64 NO { get; set; }
        public string SUPPLIER_NAME { get; set; }
        public string CREATE_DATE { get; set; }
        public string EST_FORM_ID { get; set; }
        public string CONTRACT_NAME { get; set; }
        public Int32 STATUS { get; set; }
    }
    public class EstimationForm : PLAN_ITEM
    {
        public Nullable<decimal> CUM_EST_QTY { get; set; }
        public Nullable<decimal> EST_QTY { get; set; }
        public string REMARK { get; set; }
        public Int64 EST_ITEM_ID { get; set; }
        public Nullable<decimal> EST_RATIO { get; set; }
        public Int64 NO { get; set; }
        public Nullable<decimal> CUM_RECPT_QTY { get; set; }
        public Nullable<decimal> mapQty { get; set; }
        public Nullable<decimal> Quota { get; set; }

    }
}