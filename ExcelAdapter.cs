using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace AfterServiceHelper
{
    public class ExcelAdapter : BackgroundService
    {
        private readonly ILogger<ExcelAdapter> _logger;

        public ExcelAdapter(ILogger<ExcelAdapter> logger)
        {
            _logger = logger;
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            throw new NotImplementedException();
        }
    }
}
