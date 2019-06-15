using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using ExcelTCPBindings;
using ExcelTCP;
using R_TEx1316.Actions;

namespace R_TEx1316.ExcelEventRegistration
{
    public static class RangeChangeEvents
    {
        public static event EventHandler<XlRgbColor> TopBorderChanged;

        public static void SheetChange(object Sh, Range Target)
        {
            if (Target.Borders[XlBordersIndex.xlEdgeTop].Color != XlRgbColor.rgbBlack)
            {
                //Check if workbook change history contains this border change
                //if not, border color has changed and invoke #FF0000
   
                TopBorderChanged?.Invoke(Target, Target.Borders[XlBordersIndex.xlEdgeTop].Color);
            }
        }

    }
}
