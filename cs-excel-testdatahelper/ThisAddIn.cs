using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ToolTip = System.Windows.Forms.ToolTip;
using System.Diagnostics;

namespace cs_excel_testdatahelper
{
    public partial class ThisAddIn
    {
        #region 変数
        private MainForm _formMessage = null;
        #endregion

        #region イベントハンドラ
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _formMessage = new MainForm();
            _formMessage.Show();

            this.Application.SheetSelectionChange += Application_SheetSelectionChange; ;
        }
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            if (Sh == null) return;
            if (Target == null) return;

            Excel.Worksheet sheet = Sh as Excel.Worksheet;
            if (sheet == null) return;

            string formula = Target.Formula;
            if (formula == null) return;

            string value = string.Empty;

            try
            {
                StringBuilder addressBuilder = new StringBuilder();

                string address = formula;
                address = address.Replace("=", string.Empty);
                address = address.Replace(",", string.Empty);
                address = address.Replace("$", string.Empty);
                address = address.Replace("&", string.Empty);
                address = address.Replace("+", string.Empty);
                address = address.Replace("-", string.Empty);
                address = address.Replace("*", string.Empty);
                address = address.Replace("/", string.Empty);
                address = address.Replace("%", string.Empty);

                // ""内除外
                bool start = false;
                addressBuilder.Clear();
                foreach (var s in address)
                {
                    switch (s)
                    {
                        case '"':
                            start = !start;
                            break;
                        default:
                            if (!start)
                            {
                                addressBuilder.Append(s);
                            }
                            break;
                    }

                }
                address = addressBuilder.ToString();

                // ()内取得
                int nestCount = 0;
                int nestCountMax = 0;
                addressBuilder.Clear();
                foreach (var s in address)
                {
                    switch (s)
                    {
                        case '(':
                            nestCount++;
                            nestCountMax = Math.Max(nestCountMax, nestCount);
                            if (nestCount == nestCountMax)
                            {
                                addressBuilder.Clear();
                            }
                            break;
                        case ')':
                            nestCount--;
                            break;
                        default:
                            if (nestCount == nestCountMax)
                            {
                                addressBuilder.Append(s);
                            }
                            break;
                    }
                }
                address = addressBuilder.ToString();

                // 値取得
                Excel.Range temp = sheet.get_Range(address);
                value = sheet.Cells[temp.Row, 2].Value.ToString();
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
            }

            _formMessage.Text = "[" + value + "]";

        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #endregion

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    class MainForm : Form
    {
        public MainForm()
        {
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
