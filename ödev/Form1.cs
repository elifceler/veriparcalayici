using ClosedXML.Excel;
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

namespace ödev
{
    public partial class Form1 : Form

    {
        private static string testData = "value1,value2,value3,value4,value5";
        private static int fileCount = 0;
        private static String saveDirectory = @""; //dosya yolunu gir
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
                    worksheet.Cell(1, i + 1).Value = values[i];
                }

                fileCount++;
                string fileName = Path.Combine(saveDirectory, $"file{fileCount}.xlsx");

                workbook.SaveAs(fileName);
                MessageBox.Show("File saved successfully!");
            }
        }
    }
}
