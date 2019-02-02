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
    
    public partial class PLAN_ESTIMATION_FORM
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public PLAN_ESTIMATION_FORM()
        {
            this.PLAN_ESTIMATION_HOLDPAYMENT = new HashSet<PLAN_ESTIMATION_HOLDPAYMENT>();
        }
    
        public string EST_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public string PLUS_TAX { get; set; }
        public Nullable<decimal> TAX_AMOUNT { get; set; }
        public Nullable<decimal> PAYMENT_TRANSFER { get; set; }
        public Nullable<decimal> PAYMENT_DEDUCTION { get; set; }
        public Nullable<decimal> PAID_AMOUNT { get; set; }
        public Nullable<decimal> FOREIGN_PAYMENT { get; set; }
        public Nullable<decimal> RETENTION_PAYMENT { get; set; }
        public Nullable<decimal> OTHER_PAYMENT { get; set; }
        public Nullable<decimal> PREPAY_AMOUNT { get; set; }
        public string REMARK { get; set; }
        public string CREATE_ID { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string SETTLEMENT { get; set; }
        public string TYPE { get; set; }
        public Nullable<int> STATUS { get; set; }
        public Nullable<decimal> TAX_RATIO { get; set; }
        public Nullable<System.DateTime> MODIFY_DATE { get; set; }
        public string INVOICE { get; set; }
        public string REJECT_DESC { get; set; }
        public string PROJECT_NAME { get; set; }
        public string PAYEE { get; set; }
        public Nullable<System.DateTime> PAYMENT_DATE { get; set; }
        public string INDIRECT_COST_TYPE { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PLAN_ESTIMATION_HOLDPAYMENT> PLAN_ESTIMATION_HOLDPAYMENT { get; set; }
    }
}
