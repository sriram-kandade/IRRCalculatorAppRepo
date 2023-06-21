using ClosedXML.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using ConsoleApp1.Models;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using ConsoleApp1.Utility;
using System.IO;
using ConsoleApp1.Core;

namespace ConsoleApp1
{
    public class Program
    {
        public static void Main(string[] args)
        {
            IServiceCollection services = new ServiceCollection();

            Startup startup = new Startup();
            startup.ConfigureServices(services);
            
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            var service = serviceProvider.GetService<IIRRCalculator>();
            service.BeginCalculate();
        }
    }

}

   

