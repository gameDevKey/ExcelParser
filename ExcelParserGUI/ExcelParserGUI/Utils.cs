using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelParserGUI
{
    internal class Utils
    {
        public static void SaveFile(string path, string content) {
            ShowMsg("SaveFile --> " + path + "\n" + content);
        }

        public static void ShowErrorMsg(string msg)
        {
            MessageBox.Show(msg, "发生错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void ShowMsg(string msg)
        {
            MessageBox.Show(msg, "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
