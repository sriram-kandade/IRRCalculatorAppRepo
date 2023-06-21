using ConsoleApp1.Models;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Utility
{
    public class LoanCalculator : ILoanCalculator 
    {
        public (List<Result>, double[], List<string> errors) GetResultSheetData(InputParams inputParameters, Dictionary<string, double> chargedOff, Dictionary<string, double> prepay)
        {
            List<string> errors = new List<string>();
            int term;
            double couponRate;
            double invested;
            double outstandingBalance;
            double recoveryRate;
            double purchasePremium;
            double servicingFee;
            double earnoutFee;
            double defaultMuliplier;
            double prepayMultiplier;
            DateTime valuationDate;
            DateTime issueDate;

            if (!int.TryParse(inputParameters.Term, out term))
                errors.Add("Invalid term");

            if(!double.TryParse(inputParameters.CouponRate, out couponRate))
                errors.Add("Invalid CouponRate");

            if (!double.TryParse(inputParameters.Invested, out invested))
                errors.Add("Invalid Invested");

            if (!double.TryParse(inputParameters.Outstanding_Balance, out outstandingBalance))
                errors.Add("Invalid Outstanding_Balance");

            if (!double.TryParse(inputParameters.Recovery_Rate, out recoveryRate))
                errors.Add("Invalid Recovery_Rate");

            if (!double.TryParse(inputParameters.Purchase_Premium, out purchasePremium))
                errors.Add("Invalid Purchase_Premium");

            if (!double.TryParse(inputParameters.Servicing_Fee, out servicingFee))
                errors.Add("Invalid Servicing_Fee");

            if (!double.TryParse(inputParameters.Earnout_Fee, out earnoutFee))
                errors.Add("Invalid Earnout_Fee");

            if (!double.TryParse(inputParameters.Default_Multiplier, out defaultMuliplier))
                errors.Add("Invalid Default_Multiplier");

            if (!double.TryParse(inputParameters.Prepay_Multiplier, out prepayMultiplier))
                errors.Add("Invalid Prepay_Multiplier");

            if(!DateTime.TryParse(inputParameters.Valuation_Date, out valuationDate))
                errors.Add("Invalid Valuation_Date");

            if (!DateTime.TryParse(inputParameters.Issue_Date, out issueDate))
                errors.Add("Invalid Issue_Date");

            if (errors.Count > 0) return (null, null, errors);

            string grade = inputParameters.Grade;

            double previousBalance = invested;
            double previousScheduledBalance = invested;
            double previousDefaultRate = 0;
            double[] cashFlows = new double[term + 1];

            List<Result> results = new List<Result>();

            for (int i = 0; i <= term; i++)
            {
                var months = i + 1;
                var paydate = i == 0 ? issueDate.ToString("MM/dd/yyyy") : issueDate.AddMonths(i).ToString("MM/dd/yyyy");
                var scheduledPrincipal = i == 0 ? 0 : Financial.PPmt(couponRate / 12, i, term, -1 * invested);
                var scheduledInterest = i == 0 ? 0 : Financial.Pmt(couponRate / 12, term, -1 * invested) - scheduledPrincipal;
                var prepaySpeed = 0.0;
                var defaultRate = 0.0;
                try
                {
                    prepaySpeed = i == 0 ? 0 : prepay[$"{term}M-{i}"];
                }
                catch (Exception)
                {
                    errors.Add("Prepay sheet does not contain the term specified");
                    return (null, null, errors);
                }

                try
                {
                    defaultRate = chargedOff[$"{term}-{grade}-{months}"];
                }

                catch (Exception)
                {
                    errors.Add("Charged Off sheet does not contain the term/grade specified");
                    return (null, null, errors);
                }
                var earnoutCF = months == 13 || months == 19 ? (earnoutFee / 2) * invested : 0;
                var scheduledBalance = previousScheduledBalance - scheduledPrincipal;
                var defaultVal = i == 0 ? 0 : previousBalance * defaultMuliplier * previousDefaultRate;
                var prepayVal = i == 0 ? 0 : ((previousBalance - ((previousBalance - scheduledInterest) / previousScheduledBalance) * scheduledPrincipal) * prepaySpeed) * prepayMultiplier;
                var principal = i == 0 ? 0 : (((previousBalance - defaultVal) / previousScheduledBalance) * scheduledPrincipal) + prepayVal;
                var balance = i == 0 ? previousBalance : previousBalance - (defaultVal + principal);
                var recovery = i == 0 ? 0 : defaultVal * recoveryRate;
                var servingCF = i == 0 ? 0 : (previousBalance - defaultVal) * servicingFee / 12;
                var interestAmount = i == 0 ? 0 : (previousBalance - defaultVal) * couponRate / 12;
                var totalCF = i == 0 ? -invested * (1 + purchasePremium) : principal + interestAmount + recovery - servingCF - earnoutCF;

                var result = new Result()
                {
                    Months = months,
                    Paymnt_Count = i,
                    Paydate = paydate,
                    Scheduled_Principal = string.Format("{0:#,##0.##}", scheduledPrincipal),
                    Scheduled_Interest = string.Format("{0:#,##0.##}", scheduledInterest),
                    Prepay_Speed = string.Format("{0:0.00%}", prepaySpeed),
                    Default_Rate = string.Format("{0:0.00%}", defaultRate),
                    Earnout_CF = string.Format("{0:#,##0.##}", earnoutCF),
                    Scheduled_Balance = string.Format("{0:#,##0.##}", scheduledBalance),
                    Default = string.Format("{0:#,##0.##}", defaultVal),
                    Prepay = string.Format("{0:#,##0.##}", prepayVal),
                    Principal = string.Format("{0:#,##0.##}", principal),
                    Balance = string.Format("{0:#,##0.##}", balance),
                    Recovery = string.Format("{0:#,##0.##}", recovery),
                    Servicing_CF = string.Format("{0:#,##0.##}", servicingFee),
                    Interest_Amount = string.Format("{0:#,##0.##}", interestAmount),
                    Total_CF = string.Format("{0:#,##0.##}", totalCF)
                };


                previousBalance = balance;
                previousScheduledBalance = scheduledBalance;
                previousDefaultRate = defaultRate;

                results.Add(result);
                cashFlows[i] = totalCF;
            }

            return (results, cashFlows, errors);
        }

        public double GetIRR(double[] cashFlows)
        {
            double precision = 0.00001;
            double lowRate = 0.0;
            double highRate = 1.0;

            while (Math.Abs(Npv(highRate, cashFlows)) > precision)
            {
                double currentRate = (lowRate + highRate) / 2;

                if (Npv(currentRate, cashFlows) * Npv(lowRate, cashFlows) < 0)
                    highRate = currentRate;
                else
                    lowRate = currentRate;
            }

            return (lowRate + highRate) / 2;
        }

        private double Npv(double rate, double[] cashFlows)
        {
            double npv = 0;
            for (int i = 0; i < cashFlows.Length; i++)
            {
                npv += cashFlows[i] / Math.Pow(1 + rate, i);
            }
            return npv;
        }
    }
}
