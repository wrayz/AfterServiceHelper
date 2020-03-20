using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AfterServiceHelper.Models;
using Dapper;
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
            ParseExcelObject();
            Save();

            return Task.Delay(3);
        }

        private void ParseExcelObject()
        {
            var file = new FileInfo(@"./document/Bugzilla issue20200217.xlsx");
            using (var p = new ExcelPackage(file))
            {
                var workSheet = p.Workbook.Worksheets["出貨日期"];
                // for (var rowIndex = 2; rowIndex <= 3; rowIndex++)
                for (var rowIndex = 2; rowIndex <= workSheet.Dimension.Rows; rowIndex++)
                {
                    _details.Add(new ShippedDetail
                    {
                        ProductType = workSheet.Cells[rowIndex, 1].Text,
                        RobotSn = workSheet.Cells[rowIndex, 2].Text,
                        ControlBox = workSheet.Cells[rowIndex, 3].Text,
                        StickSn = workSheet.Cells[rowIndex, 4].Text,
                        CustomerCode = workSheet.Cells[rowIndex, 5].Text,
                        ShippedCustmer = workSheet.Cells[rowIndex, 6].Text,
                        ShippedWeek = workSheet.Cells[rowIndex, 7].Text,
                        ShippedDate = workSheet.Cells[rowIndex, 8].GetValue<DateTime?>(),
                        Destination = workSheet.Cells[rowIndex, 9].Text,
                        HmiVersion = workSheet.Cells[rowIndex, 10].Text,
                        PowerBoard = workSheet.Cells[rowIndex, 11].Text,
                        DriverFW = workSheet.Cells[rowIndex, 12].Text,
                        IoVersion = workSheet.Cells[rowIndex, 13].Text,
                        RtxSn = workSheet.Cells[rowIndex, 14].Text,
                        ShippedMonth = workSheet.Cells[rowIndex, 15].GetValue<int?>(),
                        ShippedYear = workSheet.Cells[rowIndex, 16].GetValue<int?>(),
                        Remark = workSheet.Cells[rowIndex, 17].Text,
                        Hw = workSheet.Cells[rowIndex, 18].Text,
                        HwService = workSheet.Cells[rowIndex, 19].Text,
                        Country = workSheet.Cells[rowIndex, 20].Text,
                        Region = workSheet.Cells[rowIndex, 21].Text,
                        CustomerType = workSheet.Cells[rowIndex, 22].Text
                    });
                }
            }
        }

        private void Save()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder
            {
                DataSource = "127.0.0.1",
                UserID = "sa",
                Password = "p@ssw0rd",
                InitialCatalog = "AfterService",
            };

            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();
                    string insertQuery = "INSERT INTO ShippedDeatils VALUES " + 
                                        "(@ProductType, @RobotSn, @ControlBox, @StickSn, @CustomerCode, @ShippedCustmer, @ShippedWeek, "+
                                        "@ShippedDate, @Destination, @HmiVersion, @PowerBoard, @DriverFW, @IoVersion, @RtxSn, @ShippedMonth, @ShippedYear, @Remark, @Hw, @HwService, @Country, @Region, @CustomerType)";
                    connection.Execute(insertQuery, _details);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }
}
