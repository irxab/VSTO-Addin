using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using IronXL;


class ThisAddIn
 public partial class test
    {
        private void test_load(object sender, RibbonUIEventArgs e)
        {

        }
        //this method is to create a button and try to display the button
        private void testButton_clicl(object sender, RibbonControlEventArgs e)
        {
            // This executed when the button is clicked
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            
            // Perform display a message box
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
        }
    }