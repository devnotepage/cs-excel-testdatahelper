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
        #region コンストラクタ
        public MainForm()
        {
            InitializeComponent();
            this.TopMost = true;
            this.FormClosing += MainForm_FormClosing;
        }
        #endregion

        #region イベントハンドラ
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }
        #endregion

        #region オーバーライド
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_NOACTIVATE = 0x8000000;
                CreateParams p = base.CreateParams;

                if (!DesignMode)
                {
                    p.ExStyle |= (WS_EX_NOACTIVATE);
                }

                return (p);
            }
        }
        #endregion

        #region 公開関数
        public void ShowList(Dictionary<int, string> values)
        {
            this.listView1.Items.Clear();
            foreach (var value in values)
            {
                this.listView1.Items.Add(new ListViewItem(new[] { value.Key.ToString(), value.Value }));
            }
        }
        #endregion
    }
}
