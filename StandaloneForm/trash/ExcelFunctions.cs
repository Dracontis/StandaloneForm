using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StandaloneForm
{
    interface ExcelFunctions
    {
        void OpenDocument(string name);
        void CloseDocument();
        void SetValue(string range, string value);
        string GetValue(string range);
        void OpenWorksheet(int AIndex);
        void Dispose();
    }
}
