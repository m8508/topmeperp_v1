using System;
using System.Collections.Generic;
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
        public IEnumerable<TND_PROJECT_ITEM> ProjectItemInMapFP { get; set; }
        //圖算消防水資料
        public IEnumerable<TND_MAP_FW> mapFW { get; set; }
        public IEnumerable<TND_PROJECT_ITEM> ProjectItemInMapFW { get; set; }
        //圖算給排水資料
        public IEnumerable<TND_MAP_PLU> mapPLU { get; set; }
        public IEnumerable<TND_PROJECT_ITEM> ProjectItemInMapPLU { get; set; }
        //圖算弱電管線資料
        public IEnumerable<MAP_LCP_VIEW> mapLCP { get; set; }
        public IEnumerable<TND_PROJECT_ITEM> ProjectItemInMapLCP { get; set; }
        //圖算電氣管線資料
        public IEnumerable<MAP_PEP_VIEW> mapPEP { get; set; }
        public IEnumerable<TND_PROJECT_ITEM> ProjectItemInMapPEP { get; set; }
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
        public Nullable<decimal> TOTAL_BUDGET { get; set; }
        public Nullable<decimal> ITEM_COST { get; set; }
        public Nullable<decimal> ITEM_BUDGET { get; set; }
        public Nullable<decimal> TOTAL_P_COST { get; set; }
        public Nullable<decimal> COST_RATIO { get; set; }
        public Nullable<decimal> AMOUNT_BY_CODE { get; set; }
        public string SYSTEM_MAIN { get; set; }
        public string SYSTEM_SUB { get; set; }
        public Nullable<decimal> CONTRACT_PRICE { get; set; }
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
        /// 採購詢價單樣本
        /// </summary>
        public IEnumerable<PLAN_SUP_INQUIRY> planTemplateForm { get; set; }
        /// <summary>
        /// 採購供應商報價單
        /// </summary>
        public IEnumerable<PLAN_SUP_INQUIRY> planFormFromSupplier { get; set; }
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
    }
    public class ContractModels
    {
        public IEnumerable<plansummary> contractItems { get; set; }
        public IEnumerable<plansummary> wagecontractItems { get; set; }
        public IEnumerable<PLAN_PAYMENT_TERMS> paymentTerms { get; set; }
        public PLAN_PAYMENT_TERMS planpayment { get; set; }
        public IEnumerable<PLAN_ITEM> planItems { get; set; }

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
        public Int64 DELIVERY_ID { get; set; }

    }
    public class PRFunction
    {
        public Int64 NO { get; set; }
        public string TASK_NAME { get; set; }
        public string CREATE_DATE { get; set; }
        public string PR_ID { get; set; }
        public string SUPPLIER_ID { get; set; }
        public Int32 STATUS { get; set; }
    }
    public class PurchaseRequisitionDetail
    {
        public TND_PROJECT prj { get; set; }
        public PLAN_PURCHASE_REQUISITION planPR { get; set; }
        public IEnumerable<PurchaseRequisition> planPRItem { get; set; }
    }

    public class PurchaseOrderFunction
    {
        public string KEYNAME { get; set; }
        public string PR_ID { get; set; }
        public string CREATE_DATE { get; set; }
        public string SUPPLIER_ID { get; set; }
        public string NEED_DATE { get; set; }
    }
    /***
     * 施工日報表料件記錄
     */
    public class DailyReportItem
    {
        public Nullable<Int64> TASKUID { get; set; }
        public int PRJ_UID { get; set; }
        public string PROJECT_ID { get; set; }
        public string PROJECT_ITEM_ID { get; set; }
        public string ITEM_DESC { get; set; }
        public Nullable<decimal> QTY { get; set; } //圖算數量
        public Nullable<decimal> ACCUMULATE_QTY { get; set; }//累積數量
        public Nullable<decimal> FINISH_QTY { get; set; }//施作數量
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
        public Nullable<decimal> CUMULATIVE_QTY { get; set; }
        public Nullable<decimal> EST_QTY { get; set; }
        public string REMARK { get; set; }
        public Int64 EST_ITEM_ID { get; set; }
        public Nullable<decimal> EST_RATIO { get; set; }
        public Int64 NO { get; set; }
    }
    public class EstimationFormDetail
    {
        public TND_PROJECT prj { get; set; }
        public PLAN_ESTIMATION_FORM planEST { get; set; }
        public IEnumerable<EstimationForm> planESTItem { get; set; }
    }
}