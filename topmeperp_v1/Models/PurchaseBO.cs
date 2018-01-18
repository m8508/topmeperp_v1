using System;
using System.Collections.Generic;

namespace topmeperp.Models
{
    //採發階段之發包標的
    public class PURCHASE_ORDER: PLAN_SUP_INQUIRY
    {
        public Nullable<decimal> BudgetAmount { get; set; }
        public Nullable<int> CountPO { get; set; }
        public Int64 NO { get; set; }
        public string Bargain { get; set; }
    }
    public class BUDGET_SUMMANY 
    {
        public Nullable<decimal> Material_Budget { get; set; }
        public Nullable<decimal> Wage_Budget { get; set; }
    }
}