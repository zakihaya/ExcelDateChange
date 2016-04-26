using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;


namespace ExcelDataChange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                string excelName = ofd.FileName;

                Excel.Application oXls;
                Excel.Workbook oWBook;

                oXls = new Excel.Application();
                oXls.Visible = true;

                // Excelファイルをオープンする
                oWBook = (Excel.Workbook)(oXls.Workbooks.Open(
                  excelName,  // オープンするExcelファイル名
                  Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                  Type.Missing, // （省略可能）ReadOnly (True / False )
                  Type.Missing, // （省略可能）Format
                                // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                                // 5:なし / 6:引数 Delimiterで指定された文字
                  Type.Missing, // （省略可能）Password
                  Type.Missing, // （省略可能）WriteResPassword
                  Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
                  Type.Missing, // （省略可能）Origin
                  Type.Missing, // （省略可能）Delimiter
                  Type.Missing, // （省略可能）Editable
                  Type.Missing, // （省略可能）Notify
                  Type.Missing, // （省略可能）Converter
                  Type.Missing, // （省略可能）AddToMru
                  Type.Missing, // （省略可能）Local
                  Type.Missing  // （省略可能）CorruptLoad
                ));

                //Microsoft.Office.Core.DocumentProperties properties;
                //object DocProps = oWBook.BuiltinDocumentProperties;
                //properties = (Microsoft.Office.Core.DocumentProperties) oWBook.BuiltinDocumentProperties;

                //Microsoft.Office.Core.DocumentProperty prop;
                //prop = properties["Revision Number"];

                dynamic properties = oWBook.BuiltinDocumentProperties;
                dynamic property = properties.Item("Author").Value = "あいうえお";

                oWBook.Save();
                oWBook.Close();
                oXls.Quit();

                int i = 1;
            }
        }
    }
}
