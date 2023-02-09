using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        private static Timer aTimer;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            aTimer = new Timer(100000); //1.6 sec
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Enabled = true;

        }
        private static void OnTimedEvent(object source,ElapsedEventArgs e)
        {
            ExecuteService();
        }

        protected override void OnStop()
        {
            aTimer.Stop();
        }
        private static void ExecuteService()
        {
            DateTime dateTime = DateTime.Now;
            string date = dateTime.ToString("dd_mm_yyyy hh_mm");
            var file = new FileInfo(@"D:\Project\Sample.xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage(file))
            {
                ExcelWorksheet sheet = excel.Workbook.Worksheets["Sheet1"];
                string connectionString = "Data Source=RAJASHEKHAR;Initial Catalog=DemoDB;Trusted_Connection=True";
                SqlConnection connect = new SqlConnection(connectionString);
                connect.Open();
                var command = new SqlCommand("select * from dbo.Username", connect);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                int count = dataTable.Rows.Count;
                sheet.Cells.LoadFromDataTable(dataTable, true);
                FileInfo excelFile = new FileInfo(@"D:\Project\Result\Result"+date+".xlsx");
                excel.SaveAs(excelFile);
                connect.Close();
            }
        }
    }
}
