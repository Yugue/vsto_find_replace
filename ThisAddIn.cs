using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace HelloWorld
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }
               
        //Run through all the files in the directory, and find and replace
        public void find_replace(string _path)
        {
            string path = _path;
            object m = Type.Missing;

            var xlApp = new Microsoft.Office.Interop.Excel.Application();

            DirectoryInfo d = new DirectoryInfo(path);

            FileInfo[] listOfFiles_1 = d.GetFiles("*.xlsx*").ToArray();
            FileInfo[] listOfFiles_2 = d.GetFiles("*.xls*").ToArray();            
            FileInfo[] listOfFiles_3 = d.GetFiles("*.xlsm*").ToArray();
            FileInfo[] listOfFiles_4 = d.GetFiles("*.xltx*").ToArray();
            FileInfo[] listOfFiles_5 = d.GetFiles("*.xltm*").ToArray();

            FileInfo[] listOfFiles = (listOfFiles_1.Concat(listOfFiles_2)).ToArray();
            listOfFiles = (listOfFiles.Concat(listOfFiles_3)).ToArray();
            listOfFiles = (listOfFiles.Concat(listOfFiles_4)).ToArray();
            listOfFiles = (listOfFiles.Concat(listOfFiles_5)).ToArray();

            xlApp.DisplayAlerts = false;
            int count = 0;
            foreach (FileInfo file in listOfFiles)
            {

                var xlWorkBook = xlApp.Workbooks.Open(file.FullName);
                string file_name = file.FullName.Remove(file.FullName.Length - 5);
                string file_ext = file.FullName.Substring(file.FullName.Length - 4);

                foreach (Excel.Worksheet xlWorkSheet in xlWorkBook.Worksheets)
                {
                    // get the used range. 
                    Excel.Range r = (Excel.Range)xlWorkSheet.UsedRange;

                    // call the replace method to replace instances. 
                    bool success = (bool)r.Replace(
                        "Designer",
                        "Engineer",
                        Excel.XlLookAt.xlPart,
                        Excel.XlSearchOrder.xlByRows, false, m, m, m);
                    count++;
                }
                if (file_ext == "xltx")
                {
                    xlWorkBook.SaveAs(@file_name, Excel.XlFileFormat.xlOpenXMLTemplate,
                        missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
                        missing, missing, missing, missing, missing);
                }

                else if (file_ext == "xltm")
                {
                    xlWorkBook.SaveAs(@file_name, Excel.XlFileFormat.xlOpenXMLTemplateMacroEnabled,
                        missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
                        missing, missing, missing, missing, missing);
                }
                else
                {
                    xlWorkBook.Save();
                }
                xlWorkBook.Close();
            }

            MessageBox.Show("Files found: " + count.ToString());
            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
