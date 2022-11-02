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
            Random random = new Random();
            int price1 = 120000;
            int price2 = 2000;
            int amm1 = random.Next(0, 10);
            int amm2 = random.Next(0, 10);


            object start = 0, end = 0;
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Range range = doc.Range(ref start, ref end);
            Word.Table table = doc.Tables.Add(range, 3, 4, true, true);
            table.Cell(1, 1).Range.Text = "Товар";
            table.Cell(1, 2).Range.Text = "Цена";
            table.Cell(1, 3).Range.Text = "Кол-во";
            table.Cell(1, 4).Range.Text = "Сумма";

            table.Cell(2, 1).Range.Text = "Ноутбук Lenovo Legion 5";
            table.Cell(2, 2).Range.Text = price1.ToString();
            table.Cell(2, 3).Range.Text = amm1.ToString();
            table.Cell(2, 4).Range.Text = (price1*amm1).ToString();

            table.Cell(3, 1).Range.Text = "Мышка Logitech G102";
            table.Cell(3, 2).Range.Text = price2.ToString();
            table.Cell(3, 3).Range.Text = amm2.ToString();
            table.Cell(2, 4).Range.Text = (price2 * amm2).ToString();
            app.Visible = true;
        }
    }
}
