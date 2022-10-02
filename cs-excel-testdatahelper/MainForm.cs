using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace cs_excel_testdatahelper
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            this.TopMost = true;
            this.Height = 0;
            this.Width = 1000;

            this.FormClosing += MainForm_FormClosing;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        private const int WS_EX_NOACTIVATE = 0x8000000;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams p = base.CreateParams;

                if (!DesignMode)
                {
                    p.ExStyle |= (WS_EX_NOACTIVATE);
                }

                return (p);
            }
        }
    }
}
