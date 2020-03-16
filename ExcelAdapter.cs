using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AfterServiceHelper.Models;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace AfterServiceHelper
{
    public class ExcelAdapter : BackgroundService
    {
        private readonly ILogger<ExcelAdapter> _logger;
        private List<ShippedDetail> _details;

        public ExcelAdapter(ILogger<ExcelAdapter> logger)
        {
            _logger = logger;
            _details = new List<ShippedDetail>();
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            var file = new FileInfo(@"./document/Bugzilla issue20200217.xlsx");
            using(var p = new ExcelPackage(file))
            {
                var workSheet = p.Workbook.Worksheets["出貨日期"];
                for (var columnIndex = 1; columnIndex <= workSheet.Dimension.Columns; columnIndex++)
                {
                    System.Console.WriteLine($"{workSheet.Cells[1, columnIndex].Value}");
                    // for (var rowIndex = 2; rowIndex <= workSheet.Dimension.Rows; rowIndex++)
                    for (var rowIndex = 2; rowIndex <= 4; rowIndex++)
                    {
                        System.Console.WriteLine($"{workSheet.Cells[rowIndex, columnIndex].Value}");
                        _details.Add(new ShippedDetail
                        { });
                    }
                }
            }

            return Task.Delay(3);
        }
    }
}
