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
    
    public partial class PLAN_PAYMENT_TERMS
    {
        public long PAYMENT_TERMS_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public string PAYMENT_FREQUENCY { get; set; }
        public Nullable<System.DateTime> DATE_1 { get; set; }
        public Nullable<System.DateTime> DATE_2 { get; set; }
        public string PAYMENT_TERMS { get; set; }
        public Nullable<decimal> ADVANCE_RATIO { get; set; }
        public Nullable<decimal> ESTIMATED_RATIO { get; set; }
        public Nullable<decimal> RETENTION_RATIO { get; set; }
        public Nullable<decimal> USANCE { get; set; }
        public string PAYMENT { get; set; }
        public Nullable<decimal> USANCE_ADVANCE_RATIO { get; set; }
        public Nullable<decimal> USANCE_ESTIMATED_RATIO { get; set; }
        public Nullable<decimal> USANCE_RETENTION_RATIO { get; set; }
        public Nullable<decimal> PAYMENT_ADVANCE_RATIO { get; set; }
        public Nullable<decimal> PAYMENT_ESTIMATED_RATIO { get; set; }
        public Nullable<decimal> PAYMENT_RETENTION_RATIO { get; set; }
        public string CREATE_ID { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string MODIFY_ID { get; set; }
        public Nullable<System.DateTime> MODIFY_DATE { get; set; }
    }
}