using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Newtonsoft.Json;
using System.Data.Common;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

           
        }
        string pathToExcel = Environment.CurrentDirectory + @"\books.xlsx";
        string path = Environment.CurrentDirectory + @"\json.txt";
        string json;
        private void Form2_Load(object sender, EventArgs e)
        {
           

            int numberOfSheets = 0;

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(pathToExcel);// Open(pathToExcel, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            var connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""
            ", pathToExcel);

            // for finding the total number of sheets
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["TABLE_NAME"].ToString().EndsWith("$"))
                        {
                            numberOfSheets++;
                        }
                    }
                  //  MessageBox.Show(numberOfSheets.ToString());
                }
            }

            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var cmd = conn.CreateCommand();
                richTextBox1.Clear();
                //for (int i = 1; i <= numberOfSheets; i++)
                // {
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Sheets)
                {
                    //var sheetName = "Sheet"+i;
                    cmd.CommandText = String.Format(@"SELECT * FROM [{0}$]", sheet.Name);
                    
                    using (var rdr = cmd.ExecuteReader())
                    {
                        //LINQ query - when executed will create anonymous objects for each row
                        int g = 0;
                        System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        var query =
                            (from DbDataRecord row in rdr select  row).Select(x =>
                             {
                                
                                int num = dt.Rows.IndexOf(row);
                                 int j = rdr.FieldCount;
                                 int k ;
                                 g++;
                                 //dynamic item = new ExpandoObject();
                                 Dictionary<string, object> item1 = new Dictionary<string, object>();
                                 Dictionary<string, object> item = new Dictionary<string, object>();                                
                                 for (k = 0; k < j; k++)
                                 {
                                     item.Add(rdr.GetName(k), x[k]);
                                 }
                                 
                                 if(g == 1){
                                     
                                     item1.Add(sheet.Name, item);
                                     return item1;
                                 }
                                 
                                return item;
                             });
                            // item1.Add(sheet.Name,item );
                       
                        //Generates JSON from the LINQ query
                        json = JsonConvert.SerializeObject(query);

                        richTextBox1.AppendText(json);
                       
                    }
                    
                }
            }
           
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //MessageBox.Show(this.MyProperty);
           richTextBox1.SaveFile(path, RichTextBoxStreamType.PlainText);
            //System.IO.File.WriteAllText(path, json);
            MessageBox.Show("Saved Successfully");
        }

        public DataRow row { get; set; }
        public string MyProperty { get; set; }

    }
}
