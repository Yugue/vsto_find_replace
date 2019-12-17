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
using Microsoft.VisualBasic;

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
        public void find_replace()
        {
            string path = "";
            // prompt the user to select the folder path
            using(var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    path = fbd.SelectedPath;
                    System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString() + "at path: " + path, "Message");
                }
            }

            object m = Type.Missing;

            // Interaction.Inputbox is a user prompt, the user chooses the string he/she would like to replace
            string replace = Interaction.InputBox("Type the text you would like to replace", "Find text", "Default", -1, -1);
            string replacement = Interaction.InputBox("Replace that text with", "Replace text", "Default", -1, -1);
            
            var xlApp = new Microsoft.Office.Interop.Excel.Application();        

          // grab all the excel files at that directory
            DirectoryInfo d = new DirectoryInfo(@path);
            FileInfo[] listOfFiles_1 = d.GetFiles("*.xlsx*").ToArray();
            FileInfo[] listOfFiles_2 = d.GetFiles("*.xls*").ToArray();            
            FileInfo[] listOfFiles_3 = d.GetFiles("*.xlsm*").ToArray();
            FileInfo[] listOfFiles_4 = d.GetFiles("*.xltx*").ToArray();
            FileInfo[] listOfFiles_5 = d.GetFiles("*.xltm*").ToArray();

            // produce a list of files in an array
            FileInfo[] listOfFiles = (listOfFiles_1.Concat(listOfFiles_2)).ToArray();
            listOfFiles = (listOfFiles.Concat(listOfFiles_3)).ToArray();
            listOfFiles = (listOfFiles.Concat(listOfFiles_4)).ToArray();
            listOfFiles = (listOfFiles.Concat(listOfFiles_5)).ToArray();

            xlApp.DisplayAlerts = false;            
            // traverse through each file
            foreach (FileInfo file in listOfFiles)
            {

                var xlWorkBook = xlApp.Workbooks.Open(file.FullName);
                string file_name = file.FullName.Remove(file.FullName.Length - 5);
                string file_ext = file.FullName.Substring(file.FullName.Length - 4);
                // traverse through each worksheet in each file
                foreach (Excel.Worksheet xlWorkSheet in xlWorkBook.Worksheets)
                {
                    // get the used range. 
                    Excel.Range r = (Excel.Range)xlWorkSheet.UsedRange;

                    // call the replace method to replace instances. 
                    bool success = (bool)r.Replace(
                        replace,
                        replacement,
                        Excel.XlLookAt.xlPart,
                        Excel.XlSearchOrder.xlByRows, false, m, m, m);
                }
                // if file is xltx, has to save as to overwrite the new changes
                if (file_ext == "xltx")
                {
                    xlWorkBook.SaveAs(@file_name, Excel.XlFileFormat.xlOpenXMLTemplate,
                        missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
                        missing, missing, missing, missing, missing);
                }
                // if file is xltm, has to save as to overwrite the new changes
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

            MessageBox.Show("Replacement Completed");
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
