using System;

namespace AfterServiceHelper.Models
{
    // <Summary>
    // 出貨明細
    // </Summary>
    public class ShippedDetail
    {
        public string ProductType { get; set; }
        public string RobotSn { get; set; }
        public string ControlBox { get; set; }
        public string StickSn { get; set; }
        public string CustomerCode { get; set; }
        public string ShippedCustmer { get; set; }
        public string ShippedWeek { get; set; }
        public DateTime? ShippedDate { get; set; }
        public string Destination { get; set; }
        public string HmiVersion { get; set; }
        public string PowerBoard { get; set; }
        public string DriverFW { get; set; }
        public string IoVersion { get; set; }
        public string RtxSn { get; set; }
        public int? ShippedMonth { get; set; }
        public int? ShippedYear { get; set; }
        public string Remark { get; set; }
        public string Hw { get; set; }
        public string HwService { get; set; }
        public string Country { get; set; }
        public string Region { get; set; }
        public string CustomerType { get; set; }
    }
}
