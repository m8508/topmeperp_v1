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
        public IEnumerable<TND_MAP_FP> mapFP { get; set; }
        //圖算消防電資料
        public IEnumerable<TND_MAP_FW> mapFW { get; set; }
        //圖算給排水資料
        public IEnumerable<TND_MAP_PLU> mapPLU { get; set; }
        //圖算弱電管線資料
        public IEnumerable<TND_MAP_LCP> mapLCP { get; set; }
        //圖算電氣管線資料
        public IEnumerable<TND_MAP_PEP> mapPEP { get; set; }
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
        public Nullable<int> ITEM_COUNT { get; set; }
        public Nullable<decimal> BUDGET { get; set; }
    }
    public class SystemCost
    {
        /// 九宮格編碼長度 2 
        public string SYSTEM_MAIN { get; set; }
        /// 九宮格名稱
        public string SYSTEM_SUB { get; set; }
        public Nullable<decimal> MATERIAL_COST { get; set; }
        public Nullable<decimal> MAN_DAY { get; set; }
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
    /// <summary>
    /// 九宮格次九宮格編輯物件
    /// </summary>
    public class TyepManageModel
    {
        public REF_TYPE_MAIN MainType { get; set; }
        public IEnumerable<REF_TYPE_SUB> SubTypes { get; set; }
    }
}