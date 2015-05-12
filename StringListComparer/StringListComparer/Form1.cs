using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace StringListComparer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            FormBorderStyle = FormBorderStyle.FixedSingle;
            errorProvider1.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                errorProvider1.Clear();
                ValidateFields();
            }
            catch (ArgumentException exception)
            {
                MessageBox.Show(exception.Message);
                return;
            }

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;

            if (xlApp == null)
            {
                MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }



            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            try
            {
                List<string> list1 = TextBoxList1.Text.Split(textBoxDelimiter.Text[0]).ToList();
                List<string> list2 = TextBoxList2.Text.Split(textBoxDelimiter.Text[0]).ToList();

                int MaxCount = list1.Count > list2.Count ? list1.Count : list2.Count;

                for (int i = 1; i <= MaxCount; i++)
                {
                    if (i <= list1.Count)
                        ws.Cells[i, 1] = list1[i - 1];
                    if (i <= list2.Count) 
                        ws.Cells[i, 2] = list2[i - 1];
                }

                xlApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Failed to populate excel");

                xlApp.Quit();

            }
            finally
            {
                Action<object> release = x =>
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(x);
                    x = null;
                };
                release(ws);
                release(wb);
                release(xlApp);
                GC.Collect();
            }




        }

        private void ValidateFields()
        {
            bool hasErrors;
            ValidateTextBox(textBoxDelimiter, out hasErrors);
            ValidateTextBox(TextBoxList1, out hasErrors);
            ValidateTextBox(TextBoxList2, out hasErrors);

            if (hasErrors)
            {
                throw new ArgumentException("Provide all values!");
            }

        }

        private void ValidateTextBox(System.Windows.Forms.TextBox box, out bool hasErrors)
        {
            if (string.IsNullOrEmpty(box.Text))
            {
                errorProvider1.SetError(box, "Please fill the required field");
                hasErrors = true;
            }
            else
            {
                hasErrors = false;
            }
        }
    }
}
