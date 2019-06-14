using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace R_TEx1316
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void editBox5_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            UserAccessForm form = new UserAccessForm();
            form.checkedListBox1.Items.Add(new ExcelUser { FirstName = "Lucas", LastName = "Glass" }, false);
            form.checkedListBox1.Items.Add(new ExcelUser { FirstName = "Jon", LastName = "Deming" }, false);
            form.checkedListBox1.Items.Add(new ExcelUser { FirstName = "Friend Lee", LastName = "Deming" }, false);
            form.Show();
        }
    }
}
