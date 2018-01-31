﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    //定義針對特定View 所需的資料集合
    public class BusinessObjectModels
    {
    }
    //**備標階段標書基本資料(不包含圖算數量)
    public class TndProjectModels
    {
        //標單專案檔
        public TND_PROJECT tndProject { get; set; }
        //標單上品項明細資料
        public IEnumerable<TND_PROJECT_ITEM> tndProjectItem { get; set; }
        //專案任務分工
        public IEnumerable<TND_TASKASSIGN> tndTaskAssign { get; set; }
        //專案相關檔案
        public IEnumerable<TND_FILE> tndFile { get; set; }
    }
    #region 系統管理相關
    public class UserManageModels
    {
        //帳號資料
        public IEnumerable<SYS_USER> sysUsers { get; set; }
        //角色資料
        public IEnumerable<SYS_ROLE> sysRole { get; set; }
    }
    public class PrivilegeFunction : SYS_FUNCTION
    {
        public string ROLE_ID { get; set; }
    }
    #endregion
    public class MapInfoModels
    {
        //圖算消防電資料
        public IEnumerable<MAP_FP_VIEW> mapFP { get; set; }
        public IEnumerable<PLAN_ITEM> ProjectItemInMapFP { get; set; }
        //圖算消防水資料
        public IEnumerable<TND_MAP_FW> mapFW { get; set; }
        public IEnumerable<PLAN_ITEM> ProjectItemInMapFW { get; set; }
        //圖算給排水資料
        public IEnumerable<TND_MAP_PLU> mapPLU { get; set; }
        public IEnumerable<PLAN_ITEM> ProjectItemInMapPLU { get; set; }
        //圖算弱電管線資料
        public IEnumerable<MAP_LCP_VIEW> mapLCP { get; set; }
        public IEnumerable<PLAN_ITEM> ProjectItemInMapLCP { get; set; }
        //圖算電氣管線資料
        public IEnumerable<MAP_PEP_VIEW> mapPEP { get; set; }
        public IEnumerable<PLAN_ITEM> ProjectItemInMapPEP { get; set; }
        //圖算設備清單資料
        public IEnumerable<TND_MAP_DEVICE> mapDEVICE { get; set; }
    }
    public class InquiryFormModel
    {
        /// <summary>
        /// 詢價單樣本
        /// </summary>
        public IEnumerable<TND_PROJECT_FORM> tndTemplateProjectForm { get; set; }
        /// <summary>
        /// 供應商報價單
        /// </summary>
        public IEnumerable<TND_PROJECT_FORM> tndProjectFormFromSupplier { get; set; }
    }
    /// <summary>
    /// 提供詢價單、報價單所需資料結構
    /// </summary>
    public class InquiryFormDetail
    {
        public TND_PROJECT prj { get; set; }
        public TND_PROJECT_FORM prjForm { get; set; }
        public IEnumerable<TND_PROJECT_FORM_ITEM> prjFormItem { get; set; }
    }
    public class COMPARASION_DATA
    {
        //報價單編號
        public string FORM_ID { get; set; }
        //供應商名稱
        public string SUPPLIER_NAME { get; set; }
        public Nullable<decimal> TAmount { get; set; }
        public string FORM_NAME { get; set; }
    }
    public class DirectCost
    {
        /// 九宮格編碼長度 2 
        public string MAINCODE { get; set; }
        /// 九宮格名稱
        public string MAINCODE_DESC { get; set; }
        //完整地次九宮格編碼
        public Nullable<int> T_SUB_CODE { get; set; }
        //次九宮格編碼
        public string SUB_CODE { get; set; }
        //次九宮格名稱
        public string SUB_DESC { get; set; }
        public Nullable<decimal> MATERIAL_COST { get; set; }
        public Nullable<decimal> MAN_DAY { get; set; }
        public Nullable<decimal> MATERIAL_COST_INMAP { get; set; }
        public Nullable<decimal> MAN_DAY_INMAP { get; set; }
        public Nullable<int> ITEM_COUNT { get; set; }
        public Nullable<decimal> BUDGET { get; set; }
        public Nullable<decimal> TOTAL_COST { get; set; }
        public Nullable<decimal> MATERIAL_BUDGET { get; set; }
        public Nullable<decimal> ITEM_COST { get; set; }
        public Nullable<decimal> ITEM_BUDGET { get; set; }
        public Nullable<decimal> TOTAL_P_COST { get; set; }
        public Nullable<decimal> COST_RATIO { get; set; }
        public Nullable<decimal> AMOUNT_BY_CODE { get; set; }
        public string SYSTEM_MAIN { get; set; }
        public string SYSTEM_SUB { get; set; }
        public Nullable<decimal> CONTRACT_PRICE { get; set; }
        public Nullable<decimal> BUDGET_WAGE { get; set; }
        public Nullable<decimal> WAGE_BUDGET { get; set; }
        public Nullable<decimal> TOTAL_BUDGET { get; set; }
        public Nullable<decimal> ITEM_BUDGET_WAGE { get; set; }
        public Nullable<decimal> MAN_DAY_4EXCEL { get; set; }

    }
    public class SystemCost
    {
        /// 九宮格編碼長度 2 
        public string SYSTEM_MAIN { get; set; }
        /// 九宮格名稱
        public string SYSTEM_SUB { get; set; }
        public Nullable<decimal> MATERIAL_COST { get; set; }
        public Nullable<decimal> MAN_DAY { get; set; }
        public Nullable<decimal> MATERIAL_COST_INMAP { get; set; }
        public Nullable<decimal> MAN_DAY_INMAP { get; set; }
        public Nullable<int> ITEM_COUNT { get; set; }
    }
    public class SupplierFormFunction : TND_PROJECT_FORM
    {
        public Int64 NO { get; set; }
        public Nullable<decimal> TOTAL_PRICE { get; set; }
    }
    //九宮格與次九宮格索引 for 專案建立空白詢價單
    public class TYPE_CODE_INDEX
    {
        public string TYPE_CODE_1 { get; set; }
        public string TYPE_CODE_1_NAME { get; set; }
        public string TYPE_CODE_2 { get; set; }
        public string TYPE_CODE_2_NAME { get; set; }
    }
    public class PROJECT_ITEM_WITH_WAGE : TND_PROJECT_ITEM
    {
        public Nullable<decimal> MAP_QTY { get; set; }
        public Nullable<decimal> RATIO { get; set; }
        public Nullable<decimal> PRICE { get; set; }
    }
    public class PurchaseFormModel
    {
        /// <summary>
        /// 預算金額
        /// </summary>
        public BUDGET_SUMMANY BudgetSummary { get; set; }
        /// <summary>
        /// 採購詢價單樣本
        /// </summary>
        public IEnumerable<PLAN_SUP_INQUIRY> planTemplateForm { get; set; }
        /// <summary>
        /// 採購供應商報價單
        /// </summary>
        public IEnumerable<PLAN_SUP_INQUIRY> planFormFromSupplier { get; set; }
        /// <summary>
        /// 含工帶料報價單
        /// </summary>
        public IEnumerable<PLAN_SUP_INQUIRY> planForm4All { get; set; }
        /// <summary>
        /// 材料報價單樣本與預算
        /// </summary>
        public IEnumerable<PURCHASE_ORDER> materialTemplateWithBudget { get; set; }
        /// <summary>
        /// 代工報價單樣本與預算
        /// </summary>
        public IEnumerable<PURCHASE_ORDER> wageTemplateWithBudget { get; set; }

    }
    public class PlanSupplierFormFunction : PLAN_SUP_INQUIRY
    {
        public Int64 NO { get; set; }
        public Nullable<decimal> TOTAL_PRICE { get; set; }
    }
    public class PurchaseFormDetail
    {
        public TND_PROJECT prj { get; set; }
        public PLAN_SUP_INQUIRY planForm { get; set; }
        public IEnumerable<PLAN_SUP_INQUIRY_ITEM> planFormItem { get; set; }
    }
    public class COMPARASION_DATA_4PLAN
    {
        //報價單編號
        public string INQUIRY_FORM_ID { get; set; }
        //供應商名稱
        public string SUPPLIER_NAME { get; set; }
        public Nullable<decimal> TAmount { get; set; }
        public string STATUS { get; set; }
        public string FORM_NAME { get; set; }
        //預算金額
        public Nullable<decimal> BAmount { get; set; }
        //平均一日工資
        public Nullable<decimal> AvgMPrice { get; set; }
    }
    public class budgetsummary
    {
        public string TYPE_CODE_1 { get; set; }
        public string TYPE_CODE_2 { get; set; }
        public Nullable<decimal> BAmount { get; set; }
    }

    /// <summary>
    /// 九宮格次九宮格編輯物件
    /// </summary>
    public class TyepManageModel
    {
        public REF_TYPE_MAIN MainType { get; set; }
        public IEnumerable<REF_TYPE_SUB> SubTypes { get; set; }
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
    }
    public class ContractModels
    {
        public IEnumerable<plansummary> contractItems { get; set; }
        public IEnumerable<plansummary> wagecontractItems { get; set; }
        public IEnumerable<PLAN_PAYMENT_TERMS> paymentTerms { get; set; }
        public PLAN_PAYMENT_TERMS planpayment { get; set; }
        public IEnumerable<PLAN_ITEM> planItems { get; set; }
        public TND_PROJECT prj { get; set; }
        public PLAN_ESTIMATION_FORM planEST { get; set; }
        public IEnumerable<EstimationForm> planESTItem { get; set; }

    }
    public class CostForBudget
    {
        public Int64 NO { get; set; }
        public string FORM_NAME { get; set; }
        public Nullable<decimal> COST { get; set; }
        public Nullable<decimal> BUDGET { get; set; }
    }
    public class PlanRevenue : PLAN_ITEM
    {
        public Int64 NO { get; set; }
        public Nullable<decimal> PLAN_REVENUE { get; set; }
        public string CONTRACT_ID { get; set; }
    }
    public class MAP_FP_VIEW : TND_MAP_FP
    {
        public string WIRE_DESC { get; set; }
    }
    public class MAP_PEP_VIEW : TND_MAP_PEP
    {
        public string WIRE_DESC { get; set; }
        public string GROUND_WIRE_DESC { get; set; }
    }
    public class MAP_LCP_VIEW : TND_MAP_LCP
    {
        public string WIRE_DESC { get; set; }
        public string GROUND_WIRE_DESC { get; set; }
        public string PIPE_2_DESC { get; set; }
    }
    #region 供應商管理
    public class SupplierDetail
    {
        public TND_SUPPLIER sup { get; set; }
        public IEnumerable<TND_SUP_CONTACT_INFO> contactItem { get; set; }
    }
    #endregion
    public class PurchaseRequisition : PLAN_ITEM
    {
        public Nullable<decimal> MAP_QTY { get; set; }
        public Nullable<decimal> CUMULATIVE_QTY { get; set; }
        public Nullable<decimal> INVENTORY_QTY { get; set; }
        public Nullable<decimal> NEED_QTY { get; set; }
        public string REMARK { get; set; }
        public string NEED_DATE { get; set; }
        public Int64 PR_ITEM_ID { get; set; }
        public Nullable<decimal> ORDER_QTY { get; set; }
        public Nullable<decimal> RECEIPT_QTY { get; set; }
        public Nullable<decimal> ALL_RECEIPT_QTY { get; set; }
        public Nullable<decimal> DELIVERY_QTY { get; set; }
        public Int64 NO { get; set; }
        public string DELIVERY_ORDER_ID { get; set; }
        public string PARENT_PR_ID { get; set; }
        public string PR_ID { get; set; }
        public Nullable<decimal> diffQty { get; set; }
        public Nullable<decimal> RECEIPT_QTY_BY_PO { get; set; }
    }
    public class PRFunction
    {
        public Int64 NO { get; set; }
        public string TASK_NAME { get; set; }
        public string CREATE_DATE { get; set; }
        public string PR_ID { get; set; }
        public string SUPPLIER_ID { get; set; }
        public Int32 STATUS { get; set; }
        public string ALL_KEY { get; set; }
        public string MESSAGE { get; set; }
        public string MEMO { get; set; }
        public string REMARK { get; set; }
        public string RECIPIENT { get; set; }
        public string CAUTION { get; set; }
        public string CHILD_PR_ID { get; set; }
        public string PARENT_PR_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string Dminus3day { get; set; }
        public string KEY_NAME { get; set; }

    }
    public class PurchaseRequisitionDetail
    {
        public TND_PROJECT prj { get; set; }
        public PLAN_PURCHASE_REQUISITION planPR { get; set; }
        public IEnumerable<PurchaseRequisition> planPRItem { get; set; }
        public IEnumerable<PurchaseRequisition> planDOItem { get; set; }
    }

    public class PurchaseOrderFunction
    {
        public string PROJECT_ID { get; set; }
        public string PR_ID { get; set; }
        public string CREATE_DATE { get; set; }
        public string SUPPLIER_ID { get; set; }
        public string NEED_DATE { get; set; }
    }
    #region 施工日報區塊
    /***
     * 施工日報表料件記錄
     */
    public class DailyReportItem
    {
        public Nullable<Int64> TASKUID { get; set; }
        public int PRJ_UID { get; set; }
        public string PROJECT_ID { get; set; }
        public string ITEM_ID { get; set; }
        public string PROJECT_ITEM_ID { get; set; }
        public string ITEM_DESC { get; set; }
        public Nullable<decimal> QTY { get; set; } //圖算數量
        public Nullable<decimal> ACCUMULATE_QTY { get; set; }//累積數量
        public Nullable<decimal> FINISH_QTY { get; set; }//施作數量
    }
    /***
     * 標單項目彙總表
     */
    public class SummaryDailyReport
    {
        public string PROJECT_ID { get; set; }
        public string PROJECT_ITEM_ID { get; set; }
        public string ITEM_ID { get; set; }
        public string ITEM_DESC { get; set; }
        public string ITEM_UNIT { get; set; }
        public Nullable<decimal> ITEM_QUANTITY { get; set; }//標單數量
        public string TYPE_CODE_1 { get; set; }
        public string TYPE_CODE_2 { get; set; }
        public string SYSTEM_MAIN { get; set; }
        public string SYSTEM_SUB { get; set; }
        public Nullable<decimal> QTY { get; set; } //圖算數量
        public Nullable<decimal> ACCUMULATE_QTY { get; set; }//累積數量
    }
    #endregion
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
    
    public class AdvancePaymentFunction 
    {
        public Nullable<decimal> A_AMOUNT { get; set; }
        public Nullable<decimal> B_AMOUNT { get; set; }
        public Nullable<decimal> C_AMOUNT { get; set; }
        public Nullable<decimal> CUM_A_AMOUNT { get; set; }
        public Nullable<decimal> CUM_B_AMOUNT { get; set; }
        public Nullable<decimal> CUM_C_AMOUNT { get; set; }

    }
    public class PaymentDetailsFunction
    {
        public Nullable<decimal> T_ADVANCE { get; set; }
        public Nullable<decimal> CUM_T_ADVANCE { get; set; }
        public Nullable<decimal> T_OTHER { get; set; }
        public Nullable<decimal> CUM_T_OTHER { get; set; }
        public Nullable<decimal> T_RETENTION { get; set; }
        public Nullable<decimal> CUM_T_RETENTION { get; set; }
        public Nullable<decimal> T_FOREIGN { get; set; }
        public Nullable<decimal> CUM_T_FOREIGN { get; set; }
        public Nullable<decimal> SUB_AMOUNT { get; set; }
        public Nullable<decimal> CUM_SUB_AMOUNT { get; set; }
        public Nullable<decimal> PAYABLE_AMOUNT { get; set; }
        public Nullable<decimal> CUM_PAYABLE_AMOUNT { get; set; }
        public Nullable<decimal> PAID_AMOUNT { get; set; }
        public Nullable<decimal> CUM_PAID_AMOUNT { get; set; }
        public Nullable<decimal> TAX_AMOUNT { get; set; }
        public Nullable<decimal> CUM_TAX_AMOUNT { get; set; }
        public Nullable<decimal> TOTAL_TAX_AMOUNT { get; set; }
        public Nullable<decimal> TOTAL_PAID_AMOUNT { get; set; }
        public Nullable<decimal> TOTAL_PAYABLE_AMOUNT { get; set; }
        public Nullable<decimal> TOTAL_OTHER { get; set; }
        public Nullable<decimal> TOTAL_ADVANCE { get; set; }
        public Nullable<decimal> TOTAL_RETENTION { get; set; }
        public Nullable<decimal> TOTAL_SUB_AMOUNT { get; set; }
        public Nullable<decimal> TOTAL_FOREIGN { get; set; }
        public Nullable<decimal> T_REPAYMENT { get; set; }
        public Nullable<decimal> CUM_T_REPAYMENT { get; set; }
        public Nullable<decimal> T_REFUND { get; set; }
        public Nullable<decimal> CUM_T_REFUND { get; set; }
        public Nullable<decimal> TOTAL_REPAYMENT { get; set; }
        public Nullable<decimal> TOTAL_REFUND { get; set; }
    }

    public class RePaymentFunction : PLAN_OTHER_PAYMENT
    {
        public string COMPANY_NAME { get; set; }
        public string CONTRACT_NAME { get; set; }
        public Nullable<decimal> BALANCE { get; set; }
    }

    public class CashFlowFunction
    {
        public string DATE_CASHFLOW { get; set; }
        public Nullable<decimal> AMOUNT_INFLOW { get; set; }
        public Nullable<decimal> AMOUNT_OUTFLOW { get; set; }
        public Nullable<decimal> BALANCE { get; set; }
        public Nullable<decimal> RUNNING_TOTAL { get; set; }
    }
    public class ExpenseBudgetSummary : FIN_EXPENSE_ITEM
    {
        public Nullable<decimal> JAN { get; set; }
        public Nullable<decimal> FEB { get; set; }
        public Nullable<decimal> MAR { get; set; }
        public Nullable<decimal> APR { get; set; }
        public Nullable<decimal> MAY { get; set; }
        public Nullable<decimal> JUN { get; set; }
        public Nullable<decimal> JUL { get; set; }
        public Nullable<decimal> AUG { get; set; }
        public Nullable<decimal> SEP { get; set; }
        public Nullable<decimal> OCT { get; set; }
        public Nullable<decimal> NOV { get; set; }
        public Nullable<decimal> DEC { get; set; }
        public Nullable<decimal> HTOTAL { get; set; }
        public string SUBJECT_ID { get; set; }
        public string SUBJECT_NAME { get; set; }
        public string BUDGET_YEAR { get; set; }
        public Nullable<decimal> TOTAL_BUDGET { get; set; }
        public Nullable<decimal> BUDGET_AMOUNT { get; set; }
        public Nullable<decimal> MONTH_RATIO { get; set; }
        public Nullable<decimal> YEAR_RATIO { get; set; }
        public Nullable<decimal> CUM_YEAR_AMOUNT { get; set; }
        public Nullable<decimal> CUM_BUDGET { get; set; }
        public Int64 NO { get; set; }
        public Int64 SUB_NO { get; set; }
        public Nullable<decimal> TOTAL_OPERATION_EXP { get; set; }

    }
    public class OperatingExpenseModel
    {
        public ExpenseFormFunction finEXP { get; set; }
        public IEnumerable<ExpenseBudgetSummary> finEXPItem { get; set; }
        public IEnumerable<ExpenseBudgetSummary> planEXPItem { get; set; }
    }
    public class OperatingExpenseFunction: FIN_EXPENSE_FORM
    {
        public Int64 NO { get; set; }
        public string SUBJECT_NAME { get; set; }
        public string OCCURRED_DATE { get; set; }
        public string PROJECT_NAME { get; set; }

    }
    public class PlanAccountFunction : PLAN_ACCOUNT
    {
        public string PROJECT_NAME { get; set; }
        public string PAYEE { get; set; }
        public string RECORDED_DATE { get; set; }
        public string RECORDED_AMOUNT { get; set; }
        public Int64 NO { get; set; }

    }
    public class SiteBudgetModels
    {
        //第1年度工地費用預算資料
        public IEnumerable<ExpenseBudgetSummary> firstYear { get; set; }
        //第2年度工地費用預算資料
        public IEnumerable<ExpenseBudgetSummary> secondYear { get; set; }
        //第3年度工地費用預算資料
        public IEnumerable<ExpenseBudgetSummary> thirdYear { get; set; }
        //第4年度工地費用預算資料
        public IEnumerable<ExpenseBudgetSummary> fourthYear { get; set; }
        //第5年度工地費用預算資料
        public IEnumerable<ExpenseBudgetSummary> fifthYear { get; set; }
    }

    public class ExpenseFormFunction : FIN_EXPENSE_FORM
    {
        public string PROJECT_NAME { get; set; }
    }

    public class PlanItem4Map : PLAN_ITEM
    {
        public Nullable<decimal> MAP_QTY { get; set; }
        public string formName { get; set; }
        public string formId { get; set; }
    }
    public class ExpenseBudgetByMonth 
    {
        public Nullable<decimal> JAN { get; set; }
        public Nullable<decimal> FEB { get; set; }
        public Nullable<decimal> MAR { get; set; }
        public Nullable<decimal> APR { get; set; }
        public Nullable<decimal> MAY { get; set; }
        public Nullable<decimal> JUN { get; set; }
        public Nullable<decimal> JUL { get; set; }
        public Nullable<decimal> AUG { get; set; }
        public Nullable<decimal> SEP { get; set; }
        public Nullable<decimal> OCT { get; set; }
        public Nullable<decimal> NOV { get; set; }
        public Nullable<decimal> DEC { get; set; }
        public Nullable<decimal> HTOTAL { get; set; }
    }
    public class ExpensetFromOPByMonth
    {
        public Nullable<decimal> JAN { get; set; }
        public Nullable<decimal> FEB { get; set; }
        public Nullable<decimal> MAR { get; set; }
        public Nullable<decimal> APR { get; set; }
        public Nullable<decimal> MAY { get; set; }
        public Nullable<decimal> JUN { get; set; }
        public Nullable<decimal> JUL { get; set; }
        public Nullable<decimal> AUG { get; set; }
        public Nullable<decimal> SEP { get; set; }
        public Nullable<decimal> OCT { get; set; }
        public Nullable<decimal> NOV { get; set; }
        public Nullable<decimal> DEC { get; set; }
        public Nullable<decimal> HTOTAL { get; set; }
    }
    public class ExpenseBudgetModel
    {
        public IEnumerable<ExpenseBudgetSummary> summary { get; set; }
        public IEnumerable<ExpenseBudgetByMonth> budget{ get; set; }
        public IEnumerable<ExpensetFromOPByMonth> expense { get; set; }
    }
    public class ProjectList: TND_PROJECT
    {
        public string PLAN_CREATE_DATE { get; set; }
    }
    
    public class RevenueFromOwner : PLAN_VALUATION_FORM
    {
        public Int32 VACount { get; set; }
        public Int32 isVA { get; set; }
        public Int64 NO { get; set; }
        public Nullable<decimal> AR { get; set; }
        public Nullable<decimal> contractAtm { get; set; }
        public Nullable<decimal> advancePaymentBalance { get; set; }
        public Nullable<decimal> AR_UNPAID { get; set; }
        public Nullable<decimal> AR_PAID { get; set; }
        public string RECORDED_DATE { get; set; }
        


    }
}