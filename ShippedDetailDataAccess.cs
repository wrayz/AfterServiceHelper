using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using AfterServiceHelper.Models;
using Dapper;
using Microsoft.Extensions.Configuration;

namespace AfterServiceHelper
{
    //出貨明細資料存取
    class ShippedDetailDataAccess
    {
        private IConfiguration _configuration;

        public ShippedDetailDataAccess(IConfiguration configuration)
        {
            this._configuration = configuration;
        }

        public void Save(List<ShippedDetail> data)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(_configuration.GetConnectionString("DefaultConnection"));

            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                connection.Open();

                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        DeleteAll(connection, transaction);
                        InsertList(data, connection, transaction);
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw ex;
                    }
                }
            }
        }

        private static void InsertList(List<ShippedDetail> data, SqlConnection connection, SqlTransaction transaction)
        {
            var insertQuery = "INSERT INTO ShippedDeatils VALUES " +
                                        "(@ProductType, @RobotSn, @ControlBox, @StickSn, @CustomerCode, @ShippedCustmer, @ShippedWeek, " +
                                        "@ShippedDate, @Destination, @HmiVersion, @PowerBoard, @DriverFW, @IoVersion, @RtxSn, @ShippedMonth, @ShippedYear, @Remark, @Hw, @HwService, @Country, @Region, @CustomerType)";
            connection.Execute(insertQuery, data, transaction: transaction);
        }

        private void DeleteAll(SqlConnection connection, SqlTransaction transaction)
        {
            var deleteQuery = "DELETE FROM ShippedDeatils";
            connection.Execute(deleteQuery, transaction: transaction);
        }

    }
}
