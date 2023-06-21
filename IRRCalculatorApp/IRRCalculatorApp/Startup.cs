using ConsoleApp1.Core;
using ConsoleApp1.Utility;
using Microsoft.Extensions.DependencyInjection;

namespace ConsoleApp1
{
    public class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<IIRRCalculator, IRRCalculator>();
            services.AddSingleton<IExcelHelper, ExcelHelper>();
            services.AddSingleton<ILoanCalculator, LoanCalculator>();
        }
    }
}
