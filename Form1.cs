﻿using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace exportExcel_v2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String[] lines = File.ReadAllLines(@"C:\Users\Joel.Santhana\Desktop\csBeginner\basic\exportExcel\table.txt");
            String[] values;

            for (int i = 0; i < lines.Length; i++)
            {
                values = lines[i].Split('/');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim();
                }

                table.Rows.Add(row);
            }
        }
        DataTable table = new DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {
            //table.Columns.Add("ID", typeof(int));
            //table.Columns.Add("Name", typeof(String));
            //table.Columns.Add("Age", typeof(int));

            //dataGridView1.DataSource = table;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectAll();
            DataObject cpydata = dataGridView1.GetClipboardContent();
            if(cpydata != null ) Clipboard.SetDataObject(cpydata);
            Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
            xlapp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            object miseddata = System.Reflection.Missing.Value;
            workbook = xlapp.Workbooks.Add(miseddata);
            
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range xlr = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 1];
            xlr.Select();

            worksheet.PasteSpecial(xlr, Type.Missing,Type.Missing, Type.Missing,true);



        }

        private async Task mainExport()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "{fE6kbS5hk-qpNXBhMQm.uqzy9sXLrbwS804TMyqnNxrARydIp4shkZPIZfjHz.orCrc3g.Qj1JcVmQVTOYHB8BPLGTk9N-uu6IGSXMPrv2XKbTcEPqsTIohaok1un5sv}");

            try
            {
                HttpResponseMessage response = await client.GetAsync("https://api.surveymonkey.com/v3/surveys/519696776/responses/118687483059/details");
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();
                MessageBox.Show(responseBody);
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Request error:{e.Message}");
            }
        }

        private async void button1_Click_1(object sender, EventArgs e)
        {
            

            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "fE6kbS5hk-qpNXBhMQm.uqzy9sXLrbwS804TMyqnNxrARydIp4shkZPIZfjHz.orCrc3g.Qj1JcVmQVTOYHB8BPLGTk9N-uu6IGSXMPrv2XKbTcEPqsTIohaok1un5sv");
            
            try
            {
                HttpResponseMessage response = await client.GetAsync("https://api.surveymonkey.com/v3/surveys/519696776/responses/118687483059/details");
                string responseBody = await response.Content.ReadAsStringAsync();
                //MessageBox.Show(responseBody);

                JObject jsonObj = JObject.Parse(responseBody);
                var quizResults = jsonObj["quiz_results"];

                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("Correct");
                dataTable.Columns.Add("Incorrect");
                dataTable.Columns.Add("Partially Correct");
                dataTable.Columns.Add("Total Questions");
                dataTable.Columns.Add("Score");
                dataTable.Columns.Add("Total Score");

                DataRow row = dataTable.NewRow();
                row["Correct"] = quizResults["correct"];
                row["Incorrect"] = quizResults["incorrect"];
                row["Partially Correct"] = quizResults["partially_correct"];
                row["Total Questions"] = quizResults["total_questions"];
                row["Score"] = quizResults["score"];
                row["Total Score"] = quizResults["total_score"];
                dataTable.Rows.Add(row);

                dataGridView1.DataSource = dataTable;

                //File.WriteAllText("quiz_result.csv", csv);
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Request error:{ex.Message}");
            }
            

        }
    }
}
