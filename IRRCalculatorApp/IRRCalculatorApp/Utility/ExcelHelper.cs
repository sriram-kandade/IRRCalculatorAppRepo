using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ConsoleApp1.Utility
{   
    public class ExcelHelper : IExcelHelper
    {
        public Dictionary<string, string> GetInputData(IXLWorkbook workbook, string sheetName)
        {
            var worksheet = workbook.Worksheets.Where(w => w.Name == sheetName).FirstOrDefault();
            if (worksheet == null) return null;

            Dictionary<string, string> inputValues = new Dictionary<string, string>();

            foreach (var row in worksheet.RowsUsed())
            {
                var key = row.Cells().ToList()[0].Value.ToString();
                var value = row.Cells().ToList()[1].Value.ToString();

                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                    inputValues.Add(key, value);
            }

            return inputValues;
        }

        public Dictionary<string, double> GetLookupData(IXLWorkbook workbook, string sheetName)
        {            
            var worksheet = workbook.Worksheets.Where(w => w.Name == sheetName).FirstOrDefault();
            if (worksheet == null) return null;

            var firstRow = worksheet.RangeUsed().FirstRow().Cells().Skip(1).Select((v, i) => new { value = v.Value, Index = i + 1 }).ToList();
            Dictionary<string, double> dict = new Dictionary<string, double>();

            foreach (var row in worksheet.RowsUsed())
            {
                if (row.RowNumber() == 1)
                    continue;

                var yKey = (int)row.Cells().First().Value;

                for (int cell = 1; cell < row.Cells().Count() - 1; cell++)
                {
                    if (!string.IsNullOrEmpty(row.Cells().ToList()[cell].Value.ToString()))
                        dict.Add($"{firstRow.First(x => x.Index == cell).value}-{yKey}", (double)row.Cells().ToList()[cell].Value);
                }
            }

            return dict;            
        }

        public DataTable GetDataTable<T>(string columnList, List<T> results)
        {
            string columnNames = columnList; //"Months,Paymnt_Count,Paydate,Scheduled_Principal,Scheduled_Interest,Scheduled_Balance,Prepay_Speed,Default_Rate,Recovery,Servicing_CF,Earnout_CF,Balance,Principal,Default,Prepay,Interest_Amount,Total_CF";
            string[] columns = columnNames.Split(',');

            DataTable dt = new DataTable();

            foreach (var column in columns)
            {
                dt.Columns.Add(column);
            }

            foreach (var result in results)
            {
                DataRow row = dt.NewRow();
                foreach (PropertyInfo p in typeof(T).GetProperties())
                {
                    row[p.Name] = p.GetValue(result);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public double GetIRR(double[] cashFlows)
        {
            double precision = 0.00001;
            double lowRate = 0.0;
            double highRate = 1.0;

            while (Math.Abs(NPV(highRate, cashFlows)) > precision)
            {
                double currentRate = (lowRate + highRate) / 2;

                if (NPV(currentRate, cashFlows) * NPV(lowRate, cashFlows) < 0)
                    highRate = currentRate;
                else
                    lowRate = currentRate;
            }

            return (lowRate + highRate) / 2;
        }

        private double NPV(double rate, double[] cashFlows)
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
