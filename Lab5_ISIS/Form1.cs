using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using static Microsoft.Office.Core.MsoTriState;
using MSOffice = MSOfficeManager;
using Microsoft.Office.Interop.Word;
using MSOfficeManager.API;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing.Text;

namespace Lab5_ISIS
{
    public partial class Form1 : Form
    {
        private int c = 1;
        public Form1()
        {
            InitializeComponent();
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Range range = doc.Range();
            Word.Table table = doc.Tables.Add(range, 3, 4, true, true);

            for (int i = 0; i < 12; i++)
            {
                range.Text = "1111";
            }

            app.Visible = true;

        }
    }
}
