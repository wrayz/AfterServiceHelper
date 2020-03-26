using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace AfterServiceHelper
{
    public class ShippedDetailTimedHostedService : IHostedService, IDisposable
    {
        private readonly ILogger<ShippedDetailTimedHostedService> _logger;
        private readonly ShippedDetailExcelParser _parser;
        private readonly ShippedDetailDataAccess _dataAccess;
        private Timer _timer;

        public ShippedDetailTimedHostedService(ILogger<ShippedDetailTimedHostedService> logger)
        {
            _logger = logger;
             _parser = new ShippedDetailExcelParser(@"./document/Bugzilla issue20200217.xlsx");
            _dataAccess = new ShippedDetailDataAccess();
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("ShippedDetail Timed Hosted Service running.");

            _timer = new Timer(DoWork, null, TimeSpan.Zero, TimeSpan.FromSeconds(5));

            return Task.CompletedTask;
        }

        private void DoWork(object state)
        {
            var data = _parser.Parsing();

            _dataAccess.Save(data);
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("ShippedDetail Timed Hosted Service is stopping.");

            _timer?.Change(Timeout.Infinite, 0);

            return Task.CompletedTask;
        }

        public void Dispose()
        {
            _timer?.Dispose();
        }

    }
}