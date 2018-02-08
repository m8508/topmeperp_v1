using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    public class BankLoanInfo
    {
        public FIN_BANK_LOAN LoanInfo { get; set; }
        public List<FIN_LOAN_TRANACTION> LoanTransaction { get; set; }
    }
}