using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace R_TEx1316.Actions
{
    public class ChartActions
    {
        public Chart ActiveChart { get; set; }

        public ChartActions(Chart chart)
        {
            ActiveChart = chart;
        }

        public void SeriesChange(int SeriesIndex, int PointIndex)
        {

        }
    }
}
