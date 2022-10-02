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
        #region プロパティ
        public MainForm HelperForm { get; private set; } = null;
        #endregion

        #region イベントハンドラ
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            HelperForm = new MainForm();
            this.Application.SheetSelectionChange += Application_SheetSelectionChange; ;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            if (!HelperForm.Visible) return;
            if (Sh == null) return;
            if (Target == null) return;

            Excel.Worksheet sheet = Sh as Excel.Worksheet;
            if (sheet == null) return;

            string formula = Target.Formula;
            if (formula == null) return;

            var values = new Dictionary<int, string>();
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
                for (int i = 0; i < 10; i++)
                {
                    values.Add(i + 1, sheet.Cells[temp.Row, i + 1].Value?.ToString() ?? string.Empty);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
            }

            // リスト表示
            HelperForm.ShowList(values);
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
}
