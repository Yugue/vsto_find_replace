using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace HelloWorld
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorksheet();
            //currentSheet.Range["A1"].Value = "Hello World";
            //currentSheet.Columns.AutoFit();
            Globals.ThisAddIn.runFiles(@"C:\Users\ychen\Desktop\test");
            //ThisAddIn.ReplaceTextInExcelFile("test.xlsx", "Engineer", "Designer");
        }
    }
}
