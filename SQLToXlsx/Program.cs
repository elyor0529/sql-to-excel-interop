using SQLToXlsx.Properties;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SQLToXlsx
{
    internal class Program
    {
        private static void Main(string[] args)
        {

            try
            {
                //files
                var fileName = Path.Combine(Environment.CurrentDirectory, "reports", DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".xlsx");

                //db
                var connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

                //interop
                var misValue = Missing.Value;
                var xlApp = new Excel.Application();
                xlApp.Visible = false;
                var xlWorkBook = xlApp.Workbooks.Add(misValue);
                var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                var stopWatch = new Stopwatch();
                stopWatch.Start();

                using (var con = new SqlConnection(connectionString))
                {
                    con.Open();

                    using (var com = con.CreateCommand())
                    {
                        com.CommandText = Settings.Default.Query1;
                        com.CommandType = CommandType.Text;
                        com.CommandTimeout = 3600;

                        using (var reader = com.ExecuteReader())
                        {
                            var rowindex = 0;

                            while (reader.Read())
                            {
                                rowindex++;

                                //cols
                                if (rowindex == 1)
                                {
                                    for (var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                                    {
                                        xlWorkSheet.Cells[rowindex, columnIndex + 1] = reader.GetName(columnIndex);
                                    }
                                }

                                //rows
                                for (var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                                {
                                    xlWorkSheet.Cells[rowindex+1, columnIndex + 1] = reader.GetValue(columnIndex);
                                }

                            }
                        }

                        xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkBook);
                        releaseObject(xlApp);
                    }

                }

                stopWatch.Stop();
                Console.WriteLine("Running time {0:g}",stopWatch.Elapsed);

                Process.Start(fileName);

            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.Message);
            }

            Console.ReadKey();
        }


        private static void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
