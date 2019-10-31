using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MyProj.Help
{
    public class Helper
    {
        public static double GetTotal(int Amount, double Price)
        {
            return Amount * Price;
        }

        public static double GetTotalVAT(int Amount, double Price)
        {
            double VAT = 0.20;
            return Amount * Price * (VAT + 1);
        }
    }
}