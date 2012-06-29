using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;

namespace StandaloneForm
{
    class ExcelLib : ExcelFunctions
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorksheet;
        Excel.Sheets xlWorksheets;

        String filename;

        public void OpenDocument(string name)
        {
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(name);
                this.filename = name;
                xlWorksheets = xlWorkbook.Sheets;
                xlWorksheet = xlWorksheets[1] as Excel.Worksheet; // set first worksheet by default
            }
            catch (Exception e)
            {
                LogError(e.Message, "OpenDocument");
                Console.WriteLine(e.Message);
            }
        }

        public void CloseDocument()
        {
            try
            {
                xlWorkbook.Save();
                xlWorkbook.Close(false, filename, null);
                Marshal.ReleaseComObject(xlWorkbook);
                //xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception e)
            {
                LogError(e.Message, "CloseDocument");
                Console.WriteLine(e.Message);
            }
        }

        public void SetValue(string range, string value)
        {
            try
            {
                Excel.Range Rng = xlWorksheet.get_Range(range, range);
                Rng.Value = value;
                object missing = Type.Missing;
                xlWorkbook.Save();
            }
            catch (Exception e)
            {
                LogError(e.Message, "SetValue");
                Console.WriteLine(e.Message);
            }
        }

        public void OpenWorksheet(int AIndex)
        {
            try
            {
                xlWorksheet = xlWorksheets[AIndex] as Excel.Worksheet;
            }
            catch (Exception e)
            {
                LogError(e.Message, "OpenWorksheet");
                Console.WriteLine(e.Message);
            }
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public string GetValue(string range)
        {
            string value = "";
            try
            {
                Excel.Range Rng = xlWorksheet.get_Range(range, range);
                value = Rng.Value as string;
            }
            catch (Exception e)
            {
                LogError(e.Message, "GetValue");
                Console.WriteLine(e.Message);
            }
            return value;
        }

        private void LogError(string message, string function)
        {
            try
            {
                string to_file = DateTime.Now.ToString() + "\t" + "In " + function + ": " + message + "\n";
                System.IO.File.WriteAllText(@"Errors.log", to_file);
            }
            catch
            {
                Console.WriteLine("Невозможно записать в лог-файл");
            }
        }
    }
}
