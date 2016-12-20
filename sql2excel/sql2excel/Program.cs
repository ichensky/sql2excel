using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace sql2excel
{

    class SqlHelper {
        public DataTable sql2table(string sql, string connectionString) {

            using (var connection =new SqlConnection(connectionString))
            {
                var command = new SqlCommand(sql, connection);

                try
                {
                    connection.Open();

                    using (var reader = command.ExecuteReader())
                    {

                        var tb = new DataTable();
                        tb.Load(reader);

                        return tb;
                    }
                }
                catch (Exception ex)
                {
                }
            }
            return null;
        }
    }

    class ExcelHelper {
        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void gen_excel(DataTable dt, string excel_file_path) {
            FileInfo file;

            begin:;
            file = new FileInfo(excel_file_path);
            if (file.Exists)
            {
                file.Delete();
                Thread.Sleep(10);
                goto begin;
            }

            var excelColumnName = GetExcelColumnName(dt.Columns.Count);
            var colums = string.Format("A1:{0}1", excelColumnName);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //Create the worksheet
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("main");

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells[colums].AutoFilter = true;
                ws.Cells[1,1,dt.Rows.Count+1, dt.Columns.Count+1].AutoFitColumns();


                var dateColumns = from DataColumn d in dt.Columns
                                  where d.DataType == typeof(DateTime) || d.ColumnName.Contains("Date")
                                  select d.Ordinal + 1;

                foreach (var dc in dateColumns)
                {
                    ws.Cells[2, dc, dt.Rows.Count + 1, dc].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                }

                //Format the header for column 
                using (ExcelRange rng = ws.Cells[colums])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.White);  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.Black);
                }

                package.Save();
            }
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count()!=3)
            {
                Console.WriteLine("Usage: sql2excel file.sql file.xlsx connectionstring");
                return;
            }
            var file_sql = args[0];
            var file_xlsx =args[1];
            var connectionstring = args[2];

            var sql = File.ReadAllText(file_sql);
            var dt = new SqlHelper().sql2table(sql, args[2]);

            if (dt.Rows.Count>0)
            {
                new ExcelHelper().gen_excel(dt, file_xlsx);
            }
        }
    }
}
