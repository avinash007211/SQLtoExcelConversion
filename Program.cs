using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;


namespace SQL_To_Excel_Conv_Eagle_1
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {

            }
            catch (Exception ex)
            {
                Console.WriteLine("Can not open connection ! ");
            }
            ExcelExample excelExample = new ExcelExample();
            excelExample.CreateExcelFile();
            Console.ReadLine();
        }
    }

    public class ExcelExample
    {

        Application excel;
        Workbook worKbooK;
        Worksheet worksheet;
        Range range;
        Borders borders;

        public void CreateExcelFile()
        {
            try
            {
                /* The code below will make the sql connection from the database and open it for communication*/
                string connetionString = null;
                SqlConnection cnn;
                connetionString = "Data Source=BTDT2323TR1911\\SQLEXPRESS;Initial Catalog=Oxide_R1_DB;User ID=sa;Password=supervisor";
                cnn = new SqlConnection(connetionString);
                System.Data.DataTable dataTable = new System.Data.DataTable();
                cnn.Open();
                SqlCommand cmd = new SqlCommand("SELECT TOP (384) *FROM FloatTable ORDER BY DateAndTime DESC,TagIndex asc; ", cnn);
                Console.WriteLine("Connnection open state :" + cnn.State.ToString());

                /* The code below will make the sql reader which will read the values from the database*/

                var reader = cmd.ExecuteReader();
                List<FloatTable> floatTables = new List<FloatTable>();

                while (reader.Read())
                {
                    FloatTable floatTable = new FloatTable();
                    string date = reader["DateAndTime"].ToString();
                    floatTable.DateAndTime = Convert.ToDateTime(date);
                    floatTable.Val = Convert.ToDecimal(reader["Val"].ToString());
                    floatTables.Add(floatTable);
                }


                cnn.Close();

                excel = new Application();
                worKbooK = excel.Workbooks.Add(Type.Missing);
                worKbooK = excel.Workbooks.Open(@"C:\Report Generartion Project\Report Formats\Oxide_Mill_Log_Sheet_Plant_1.xlsx");
                worksheet = (Worksheet)worKbooK.ActiveSheet;

                var ft = floatTables.GroupBy(o => o.DateAndTime.ToString("HH:mm:ss")); // Added "HH:mm:ss" to group by time only ine the first column 
                int rowStart = 6;
                int columnStart = 1;
                int i = 0;

                /* The code below will make the map the values read the values from the database into the excel workbook and worksheets*/

                foreach (var f in ft)
                {
                    var key = f.Key;
                    var curRow = rowStart + i;
                    worksheet.Cells[curRow, columnStart] = key;
                    var values = f.Select(o => o.Val).ToList();

                    range = worksheet.get_Range("A2", "I" + (i + 2).ToString()); //set the properties of each column to auto fit the largest cell in each col
                    range.Columns.AutoFit();

                    for (int j = 0; j < values.Count; j++)

                    {

                        worksheet.Cells[rowStart + i, columnStart + j + 1] = values[j];

                    }

                    i++;
                }

                /* The code below will save the value at the desired location*/
                var dateString = DateTime.Now.Ticks.ToString();
                excel.Application.DisplayAlerts = false;
                string destPath = "C:\\Report Generartion Project\\Reports Generated\\Oxide_R1_DB_Report.xlsx";
                worKbooK.SaveAs(destPath);
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private string GetRandomNumber()
        {
            return Guid.NewGuid().ToString();
        }

        public class FloatTable
        {
            public DateTime DateAndTime { get; set; }
            public decimal Val { get; set; }
            public bool DisplayAlerts { get; set; }

        }

    }
}
