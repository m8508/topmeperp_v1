using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    public class BusinessObject4Expense
    {
    }
    //提供費用申請流程清單所需物件
    public class ExpenseFlowTask: WF_PORCESS_TASK
    {
        //FIN_EXPENSE_FORM
        public string EXP_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public Int32 OCCURRED_YEAR { get; set; }
        public Int32 OCCURRED_MONTH { get; set; }
        public string PAYEE { get; set; }
        public string PAYMENT_DATE { get; set; }
        public string REQ_DESC { get; set; }
        //WF_PROCESS_REQUEST
        public string REQ_USER_ID { get; set; }
        public Int64 CURENT_STATE { get; set; }
        public Int64 PID { get; set; } //可關聯至WF_PROCESS
        // FROM URL
        public string FORM_URL { get; set; }
    }
}