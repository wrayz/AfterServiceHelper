using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using AfterServiceHelper.Models;
using Dapper;

namespace AfterServiceHelper
{
    //出貨明細資料存取
    class ShippedDetailDataAccess
    {
        private readonly List<ShippedDetail> _data;

        public ShippedDetailDataAccess(List<ShippedDetail> data)
        {
            _data = data;
        }

        private void Save()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder
            {
                DataSource = "192.168.134.54",
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
                                        "(@ProductType, @RobotSn, @ControlBox, @StickSn, @CustomerCode, @ShippedCustmer, @ShippedWeek, " +
                                        "@ShippedDate, @Destination, @HmiVersion, @PowerBoard, @DriverFW, @IoVersion, @RtxSn, @ShippedMonth, @ShippedYear, @Remark, @Hw, @HwService, @Country, @Region, @CustomerType)";
                    connection.Execute(insertQuery, _data);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }
}