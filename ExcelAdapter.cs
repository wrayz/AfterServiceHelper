using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace AfterServiceHelper
{
    public class ExcelAdapter : BackgroundService
    {
        private readonly ILogger<ExcelAdapter> _logger;
        private readonly ShippedDetailExcelParser _parser;
        private readonly ShippedDetailDataAccess _dataAccess;

        public ExcelAdapter(ILogger<ExcelAdapter> logger)
        {
            _logger = logger;
            _parser = new ShippedDetailExcelParser(@"./document/Bugzilla issue20200217.xlsx");
            _dataAccess = new ShippedDetailDataAccess();
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            var data = _parser.Parsing();

            _dataAccess.Save(data);

            return Task.Delay(3);
        }
    }
}
