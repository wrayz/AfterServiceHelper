using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace AfterServiceHelper
{
    public class ShippedDetailTimedHostedService : IHostedService, IDisposable
    {
        public IConfiguration _configuration { get; }

        private readonly ILogger<ShippedDetailTimedHostedService> _logger;
        private readonly ShippedDetailExcelParser _parser;
        private readonly ShippedDetailDataAccess _dataAccess;
        private readonly int _interval;
        private Timer _timer;

        public ShippedDetailTimedHostedService(ILogger<ShippedDetailTimedHostedService> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
            _parser = new ShippedDetailExcelParser(_configuration.GetValue<string>("ExcelPath"));
            _dataAccess = new ShippedDetailDataAccess(configuration);

            _interval = _configuration.GetValue<int>("CycleHour");
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("===================Running=====================");

            _timer = new Timer(DoWork, null, TimeSpan.Zero, TimeSpan.FromHours(_interval));

            return Task.CompletedTask;
        }

        private void DoWork(object state)
        {
            _logger.LogInformation($"{DateTime.Now} Parsing process");

            var data = _parser.Parsing();

            _logger.LogInformation($"{DateTime.Now} Saving into database");

            _dataAccess.Save(data);

            _logger.LogInformation($"{DateTime.Now} Save Finished");
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("===================Stopping=====================");

            _timer?.Change(Timeout.Infinite, 0);

            return Task.CompletedTask;
        }

        public void Dispose()
        {
            _timer?.Dispose();
        }

    }
}
