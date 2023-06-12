using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelParserGUI
{
    public partial class ExcelParser : Form
    {
        private const string xlsFilter = "Excel文件|*.xls;*.xlsx";

        public ExcelParser()
        {
            InitializeComponent();
            this.Text = "ExcelParser";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();       //获得路径
            onReadExcel(path);
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;                                                              //重要代码：表明是所有类型的数据，比如文件路径
            else
                e.Effect = DragDropEffects.None;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = xlsFilter;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                onReadExcel(openFileDialog1.FileName);
            }
        }

        private void onReadExcel(string filePath)
        {
            setPathTips(filePath);
            ExcelParserCore.Parse(filePath,"json");
        }

        private void setPathTips(string path)
        {
            pathTips.Text = "当前Excel:"+path;
        }
    }

}
