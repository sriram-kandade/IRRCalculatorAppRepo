using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Models
{
    public class Result
    {
        public int Months { get; set; }
        public int Paymnt_Count { get; set; }
        public string Paydate { get; set; }
        public string Scheduled_Principal { get; set; }
        public string Scheduled_Interest { get; set; }
        public string Scheduled_Balance { get; set; }
        public string Prepay_Speed { get; set; }
        public string Default_Rate { get; set; }
        public string Earnout_CF { get; set; }
        public string Balance { get; set; }
        public string Default { get; set; }
        public string Principal { get; set; }
        public string Prepay { get; set; }
        public string Recovery { get; set; }
        public string Servicing_CF { get; set; }
        public string Interest_Amount { get; set; }
        public string Total_CF { get; set; }
    }   
}
