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
    
    public partial class TND_PROJECT
    {
        public string PROJECT_ID { get; set; }
        public string PROJECT_NAME { get; set; }
        public string OWNER_USER_ID { get; set; }
        public string EXCEL_FILE_NAME { get; set; }
        public Nullable<int> START_ROW_NO { get; set; }
        public string STATUS { get; set; }
        public string MODIFY_USER_ID { get; set; }
        public Nullable<System.DateTime> MODIFY_DATE { get; set; }
        public string CREATE_USER_ID { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string CONTRUCTION_NAME { get; set; }
        public string CONTRUCTION_ADDRESS { get; set; }
        public string CUSTOMER_NAME { get; set; }
        public string CONTACT_PHONE { get; set; }
        public Nullable<System.DateTime> RECEIVED_DATE { get; set; }
        public Nullable<System.DateTime> TENDER_DATE { get; set; }
    }
}
