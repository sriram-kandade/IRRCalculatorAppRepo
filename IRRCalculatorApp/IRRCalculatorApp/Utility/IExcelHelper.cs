using ClosedXML.Excel;
using System.Collections.Generic;
using System.Data;

namespace ConsoleApp1.Utility
{
    public interface IExcelHelper
    {
        Dictionary<string, string> GetInputData(IXLWorkbook workbook, string sheetName);
        Dictionary<string, double> GetLookupData(IXLWorkbook workbook, string sheetName);
        DataTable GetDataTable<T>(string columnList, List<T> results);
        double GetIRR(double[] cashFlows);
    }
}