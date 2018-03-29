//------------------------------------------------------------------------------
// <auto-generated>
//     這個程式碼是由範本產生。
//
//     對這個檔案進行手動變更可能導致您的應用程式產生未預期的行為。
//     如果重新產生程式碼，將會覆寫對這個檔案的手動變更。
// </auto-generated>
//------------------------------------------------------------------------------

namespace topmeperp.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class FIN_BANK_LOAN
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public FIN_BANK_LOAN()
        {
            this.FIN_LOAN_TRANACTION = new HashSet<FIN_LOAN_TRANACTION>();
        }
    
        public long BL_ID { get; set; }
        public string BANK_ID { get; set; }
        public string BANK_NAME { get; set; }
        public string BRANCH_NAME { get; set; }
        public string ACCOUNT_NO { get; set; }
        public string ACCOUNT_NAME { get; set; }
        public string REF_ACCOUNT_CODE { get; set; }
        public Nullable<System.DateTime> START_DATE { get; set; }
        public Nullable<System.DateTime> DUE_DATE { get; set; }
        public Nullable<int> PERIOD_COUNT { get; set; }
        public Nullable<decimal> QUOTA { get; set; }
        public string REMARK { get; set; }
        public string CREATE_ID { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string MODIFY_ID { get; set; }
        public Nullable<System.DateTime> MODIFY_DATE { get; set; }
        public string PROJECT_ID { get; set; }
        public Nullable<decimal> AR_PAYBACK_RATIO { get; set; }
        public Nullable<decimal> CUM_AR_RATIO { get; set; }
        public Nullable<decimal> QUOTA_AVAILABLE_RATIO { get; set; }
        public string QUOTA_RECYCLABLE { get; set; }
        public string IS_SUPPLIER { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FIN_LOAN_TRANACTION> FIN_LOAN_TRANACTION { get; set; }
    }
}
