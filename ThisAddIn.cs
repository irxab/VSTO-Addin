using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ThisAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // This method is executed when the button is clicked
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            
            // Perform some action, for example, display a message box
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
            if (activeSheet != null)
            {
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;
                if (selectedRange != null)
                {
                    int rowCount = selectedRange.Rows.Count;
                    int columnCount = selectedRange.Columns.Count;

                    string message = $"Selected Range: {rowCount} rows, {columnCount} columns";
                    System.Windows.Forms.MessageBox.Show(message, "Selected Range Info");
                }
            }
        }
    }
}
