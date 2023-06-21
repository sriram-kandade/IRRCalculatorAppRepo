using ClosedXML.Excel;
using ConsoleApp1.Models;
using ConsoleApp1.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ConsoleApp1.Core
{
    public class IRRCalculator : IIRRCalculator
    {
        private readonly IExcelHelper _helper;
        private readonly ILoanCalculator _loanCalculator;

        private const string _columnList = "Months,Paymnt_Count,Paydate,Scheduled_Principal,Scheduled_Interest,Scheduled_Balance,Prepay_Speed,Default_Rate,Recovery,Servicing_CF,Earnout_CF,Balance,Principal,Default,Prepay,Interest_Amount,Total_CF";
        public IRRCalculator(IExcelHelper helper, ILoanCalculator loanCalculator)
        {
            _helper = helper;
            _loanCalculator = loanCalculator;
        }
        public void BeginCalculate()
        {
            while (true)
            {
                Console.WriteLine("Please provide the input excel sheet path........");
                var path = Console.ReadLine();


                var inputParameters = new InputParams();
                var chargedOff = new Dictionary<string, double>();
                var prepay = new Dictionary<string, double>();

                if (!File.Exists(path))
                {
                    Console.WriteLine("Invalid file path \n");
                    continue;
                }

                if (Path.GetExtension(path) != ".xlsx" && Path.GetExtension(path) != ".xls")
                {
                    Console.WriteLine("Invalid file. Please use excel file \n");
                    continue;
                }

                Console.WriteLine("\nPlease wait while we calculate the IRR....");
                using (IXLWorkbook workbook = new XLWorkbook(path))
                {
                    Dictionary<string, string> inputData = _helper.GetInputData(workbook, "Input");

                    if (inputData == null) Console.WriteLine("'Input' sheet is not present in the supplied file");

                    foreach (PropertyInfo p in typeof(InputParams).GetProperties())
                    {
                        if(inputData.ContainsKey(p.Name))
                            p.SetValue(inputParameters, inputData[p.Name]);
                    }

                    chargedOff = _helper.GetLookupData(workbook, "Charged Off");
                    if (chargedOff == null) Console.WriteLine("'Charged Off' sheet is not present in the supplied file");

                    prepay = _helper.GetLookupData(workbook, "Prepay");
                    if (prepay == null) Console.WriteLine("'Charged Off' sheet is not present in the supplied file");
                }

                (List<Result>, double[], List<string>) result = _loanCalculator.GetResultSheetData(inputParameters, chargedOff, prepay);

                if (result.Item1 == null && result.Item2 == null && result.Item3.Count > 0)
                {
                    Console.WriteLine($"Calculation failed due to the following reasons:\n {string.Join('\n', result.Item3.Select(e => $"{e}"))}");
                    continue;
                }

                var resultSheetData = _helper.GetDataTable(_columnList, result.Item1);

                string fileName = $"{Directory.GetParent(path)}\\IRRCalculationResults_{DateTime.Now.ToString("yyyyMMddmmss")}.xlsx";
                using (IXLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(resultSheetData, "IRR Calculation");
                    wb.SaveAs(fileName);
                }

                var irr = _loanCalculator.GetIRR(result.Item2) * 12;

                Console.WriteLine($"\nThe result file is saved to the following location:\n {fileName}");
                Console.WriteLine($"\nThe calculated IRR is: {string.Format("{0:0.000000000%}", irr)}");
                Console.WriteLine("\nHit Enter to start over");
                Console.ReadLine();

                continue;
            }
        }
    }
}
