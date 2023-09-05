using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

class ThisAddIn
{
    Excel.ThisAddIn newThisAddIn = this.Application.ThisAddIn.Add(System.Type.Missing)
}
