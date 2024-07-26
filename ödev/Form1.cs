using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ödev
{
    public partial class Form1 : Form

    {
        private static string testData = "value1,value2,value3,value4,value5";
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void buttonStart_Click(object sender, EventArgs e)
        {
            textBoxInput.Text = testData;
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            string[] values = textBoxInput.Text.Split(',');

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                for (int i = 0; i < values.Length; i++)
                {
                    worksheet.Cell(i + 1, 1).Value = values[i];
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Save an Excel File"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("File saved successfully!");
                }
            }
        }
    }
}
