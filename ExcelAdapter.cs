using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace AfterServiceHelper
{
    public class ExcelAdapter : BackgroundService
    {
        private readonly ILogger<ExcelAdapter> _logger;

        public ExcelAdapter(ILogger<ExcelAdapter> logger)
        {
            _logger = logger;
            System.Console.WriteLine("HIHI");
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            System.Console.WriteLine("Start");
            var fi = new FileInfo(@"./document/Bugzilla issue20200217.xlsx");
            using(var p = new ExcelPackage(fi))
            {
                var ws = p.Workbook.Worksheets["出貨日期"];
                Console.WriteLine($"Cell A1 Value: {ws.Cells.Columns}");
            }

            return Task.Delay(3);
        }
    }
}
