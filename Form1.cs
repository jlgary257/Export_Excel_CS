using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
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
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(String));
            table.Columns.Add("Age", typeof(int));

            dataGridView1.DataSource = table;
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
    }
}
