using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace R_TEx1316
{
    public partial class ThisAddIn
    {
        public Workbook ActiveWorkbook
        {
            get { return Application.ActiveWorkbook; }
        }

        public Worksheet ActiveWorksheet
        {
            get { return (Worksheet)Application.ActiveSheet; }
        }

        public Range ActiveRange { get { return activeRange ?? ActiveWorksheet?.Range["A1"]; } set { activeRange = value; } }
        private Range activeRange;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Debug.WriteLine("Starting up");
            
        }

        public void StartCollab()
        {
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.ProtectedViewWindowOpen += Application_ProtectedWorkbookOpen;
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            Application.WorkbookNewChart += Application_WorkbookNewChart;
            Application.SheetChange += Application_SheetChange;
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.SheetActivate += Application_SheetActivate;
            Application.SheetDeactivate += Application_SheetDeactivate;
        }


        private Range lastRange = null;

        private XlRgbColor lastLeftColor, lastTopColor, lastRightColor, lastBottomColor;
        private XlBorderWeight lastLeftWeight, lastTopWeight, lastRightWeight, lastBottomWeight;
        private XlLineStyle lastLeftStyle, lastTopStyle, lastRightStyle, lastBottomStyle;
        private List<string> lastComments;
        private void Application_SheetSelectionChange(object Sh, Range Target)
        {
            if (lastRange != null)
            {
                Debug.WriteLine(lastLeftColor);
                lastRange.Borders.Item[XlBordersIndex.xlEdgeLeft].Color = lastLeftColor;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeLeft].Weight = lastLeftWeight;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeLeft].LineStyle = lastLeftStyle;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeTop].Color = lastTopColor;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeTop].Weight = lastTopWeight;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeTop].LineStyle = lastTopStyle;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeRight].Color = lastRightColor;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeRight].Weight = lastRightWeight;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeRight].LineStyle = lastRightStyle;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeBottom].Color = lastBottomColor;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeBottom].Weight = lastBottomWeight;
                lastRange.Borders.Item[XlBordersIndex.xlEdgeBottom].LineStyle = lastBottomStyle;
                int count = 0;
                foreach (Range r in lastRange)
                {
                    r.ClearComments();
                    if (lastComments[count] != null)
                        r.AddComment(lastComments[count]);
                    count++;
                }
                //all you have to do is .copy!!!!!!!!!!!!!!!!!!!!!!
                Target.Copy(lastRange);
                ActiveWorksheet.Range["A1"].Value = "test";
            }
            lastRange = Target;
            Debug.WriteLine("got selection change");
            ActiveRange = Target;
            lastComments = new List<string>();

            lastLeftColor = (XlRgbColor)Target.Borders.Item[XlBordersIndex.xlEdgeLeft].Color;
            lastLeftWeight = (XlBorderWeight)Target.Borders.Item[XlBordersIndex.xlEdgeLeft].Weight;
            lastLeftStyle = (XlLineStyle)Target.Borders.Item[XlBordersIndex.xlEdgeLeft].LineStyle;
            lastTopColor = (XlRgbColor)Target.Borders.Item[XlBordersIndex.xlEdgeTop].Color;
            lastTopWeight = (XlBorderWeight)Target.Borders.Item[XlBordersIndex.xlEdgeTop].Weight;
            lastTopStyle = (XlLineStyle)Target.Borders.Item[XlBordersIndex.xlEdgeTop].LineStyle;
            lastRightColor = (XlRgbColor)Target.Borders.Item[XlBordersIndex.xlEdgeRight].Color;
            lastRightWeight = (XlBorderWeight)Target.Borders.Item[XlBordersIndex.xlEdgeRight].Weight;
            lastRightStyle = (XlLineStyle)Target.Borders.Item[XlBordersIndex.xlEdgeRight].LineStyle;
            lastBottomColor = (XlRgbColor)Target.Borders.Item[XlBordersIndex.xlEdgeBottom].Color;
            lastBottomWeight = (XlBorderWeight)Target.Borders.Item[XlBordersIndex.xlEdgeBottom].Weight;
            lastBottomStyle = (XlLineStyle)Target.Borders.Item[XlBordersIndex.xlEdgeBottom].LineStyle;

            Target.Borders.Item[XlBordersIndex.xlEdgeLeft].Color = XlRgbColor.rgbBlue;
            Target.Borders.Item[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            Target.Borders.Item[XlBordersIndex.xlEdgeTop].Color = XlRgbColor.rgbBlue;
            Target.Borders.Item[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            Target.Borders.Item[XlBordersIndex.xlEdgeRight].Color = XlRgbColor.rgbBlue;
            Target.Borders.Item[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
            Target.Borders.Item[XlBordersIndex.xlEdgeBottom].Color = XlRgbColor.rgbBlue;
            Target.Borders.Item[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            foreach (Range r in Target)
            {
                lastComments.Add(r.Comment?.Text());
                r.ClearComments();
                r.AddComment(ActiveWorkbook?.Author + ": Updating this cell at the moment");
            }



        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
            Application.WorkbookNewChart -= Application_WorkbookNewChart;
            Application.SheetChange -= Application_SheetChange;
            Application.SheetActivate -= Application_SheetActivate;
            Application.SheetDeactivate -= Application_SheetDeactivate;
        }

        private void Application_SheetDeactivate(object Sh)
        {
            Worksheet sheet = (Worksheet)Sh;
            sheet.SelectionChange -= Sheet_SelectionChange;
        }

        private void Application_SheetActivate(object Sh)
        {
            Debug.WriteLine("Sheet Activated");
            Worksheet sheet = (Worksheet)Sh;
            //sheet.SelectionChange += Sheet_SelectionChange;
        }

        private void Sheet_SelectionChange(Range Target)
        {
            Debug.WriteLine("got selection change");
            ActiveRange = Target;
            ActiveWorksheet.Range[Target].Value = "Clicked here buddy";

        }

        private void ThisAddIn_Change(Range Target)
        {

        }

        private void Application_SheetChange(object Sh, Range Target)
        {
            Debug.WriteLine("Application Sheet Change");
            ActiveWorksheet.Range["A1"].ClearComments();
            ActiveWorksheet.Range["A1"].AddComment("Generic Comment");

        }

        private void Application_WorkbookNewChart(Workbook Wb, Chart Ch)
        {

        }

        private void Application_WorkbookNewSheet(Workbook Wb, object Sh)
        {

        }

        private void Application_WorkbookOpen(Workbook Wb)
        {
            Debug.WriteLine("New workbook opened: " + Wb.Name);
        }

        private void Application_ProtectedWorkbookOpen(ProtectedViewWindow window)
        {
            Debug.WriteLine("New workbook opened: " + window.Workbook.Name);
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
