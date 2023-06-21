using ConsoleApp1.Models;
using System.Collections.Generic;

namespace ConsoleApp1.Utility
{
    public interface ILoanCalculator
    {
        (List<Result>, double[], List<string> errors) GetResultSheetData(InputParams inputParameters, Dictionary<string, double> chargedOff, Dictionary<string, double> prepay);
        double GetIRR(double[] cashFlows);
    }
}