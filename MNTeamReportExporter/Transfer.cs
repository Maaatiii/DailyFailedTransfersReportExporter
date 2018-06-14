using LinqToExcel.Attributes;
using System;

namespace MNTeamReportExporter
{
    public class Transfer
    {
        [ExcelColumn("Correlation ID")]
        public string CorellationId { get; set; }

        [ExcelColumn("Custom Data Transfer Type")]
        public string TransferType { get; set; }

        [ExcelColumn("Source Name")]
        public string SourceName { get; set; }

        [ExcelColumn("Timestamp")]
        public DateTime Timestamp { get; set; }

        [ExcelColumn("Custom Data Machine Name")]
        public string MachineName { get; set; }

        [ExcelColumn("Custom Data Version Number")]
        public string Version { get; set; }

        [ExcelColumn("Host Name")]
        public string HostName { get; set; }

        [ExcelColumn("Comment")]
        public string Comment { get; set; }

	    [ExcelColumn("JIRA")]
	    public string Jira { get; set; }

		public string TenantNo
        {
            get
            {
                return HostName.Replace("VM-","").Substring(0, 4);
            }
        }
    }
}
