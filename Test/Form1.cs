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

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            var excel = new Spire.Xls.Workbook();
            excel.LoadFromFile("../../test.xlsx");
            Stream stream = new MemoryStream();
            excel.SaveToStream(stream);
            Excel_PDF.ExcelToPDF excelToPDF = new Excel_PDF.ExcelToPDF(stream);
        }
    }
}
