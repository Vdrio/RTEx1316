using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelTCPBindings;
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
            ExcelUser me = new ExcelUser("Lucas", "Glass", UserAccessLevel.Admin);
            form.checkedListBox1.Items.Add(me);
            form.checkedListBox1.Items.Add(new ExcelUser() { FirstName = "Jon", LastName = "Deming" }, false);
            form.checkedListBox1.Items.Add(new ExcelUser() { FirstName = "Friend Lee", LastName = "Deming" }, false);
            //me.OnAccessLevelChanged += Me_OnAccessLevelChanged;
            me.AccessLevel = UserAccessLevel.Default;
            //form.checkedListBox1.ItemCheck += CheckedListBox1_ItemCheck;
            form.Show();
        }

        public void Me_OnAccessLevelChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("Ribbon detected access level change");
        }

        private void CheckedListBox1_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
        {
           CheckedListBox box = (CheckedListBox)sender;
           Debug.WriteLine(box.Items[e.Index]);
        }
    }
}
