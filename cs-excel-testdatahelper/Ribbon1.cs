using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace cs_excel_testdatahelper
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonEnable_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.HelperForm.Show();
            Globals.ThisAddIn.HelperForm.Visible = true;
        }

        private void buttonDisable_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.HelperForm.Visible = false;
        }
    }
}
