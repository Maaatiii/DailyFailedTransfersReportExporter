using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNTeamReportExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var date = DateTime.Now.Date.AddDays(-1);

            var excel = new ExcelQueryFactory(@"C:\Users\Mati\Downloads\MN_TEAM_FailedTransfersAnalysis.xlsx");

            var transfers = (from c in excel.Worksheet<Transfer>("DATA")
                             orderby c.Timestamp
                             where c.Timestamp > date && c.Timestamp < date.AddHours(24)
                            select c).ToList();

            var groupedByTrasferType = transfers.GroupBy(x => x.TransferType, x=>x,
                (key, g) => new { TransferType = key, Transfers = g.ToList() });

            foreach(var item in groupedByTrasferType)
            {
                Console.WriteLine(item.TransferType);

                var grouppedByTenant = item.Transfers.GroupBy(x => x.SourceName + " " + x.TenantNo, x=>x,
                     (key, g) => new { Tenant = key, Transfers = g.ToList() });

                foreach(var transfersForTenant in grouppedByTenant)
                {
                    Console.WriteLine(transfersForTenant.Tenant);

                    foreach (var transfer in transfersForTenant.Transfers)
                    {
                        var line = $"{transfer.Timestamp} CorellationId: {transfer.CorellationId} Machine Name: {transfer.MachineName} Version: {transfer.Version} Reason: {transfer.Comment}";

                        Console.WriteLine(line);
                    }
                }
            }

            Console.ReadKey();
        }
    }
}
