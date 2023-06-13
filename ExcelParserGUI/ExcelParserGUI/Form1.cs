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
        private string[] exportTypeList = { "json","lua" };
        private int exportTypeIndex = 0;
        private string sourcePath = string.Empty;
        private string exportPath = string.Empty;

        public ExcelParser()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //初始化comboBox1
            comboBox1.Items.Clear();
            foreach (var tpe in exportTypeList)
            {
                comboBox1.Items.Add(tpe);
            }
            comboBox1.SelectedIndex = 0;
            exportTypeIndex = comboBox1.SelectedIndex;

            pathTips.Text = "(当前路径:空)";
            exportPathTips.Text = "(导出路径:未选择)";
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();       //获得路径
            setSourcePath(path);
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
                setSourcePath(openFileDialog1.FileName);
            }
        }

        private void setSourcePath(string path)
        {
            sourcePath = path;
            pathTips.Text = string.Format("(当前路径:{0})",sourcePath);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            exportTypeIndex = comboBox1.SelectedIndex;
        }

        private void exportPathBtn_Click(object sender, EventArgs e)
        {
            var folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                exportPath = folder.SelectedPath;
                exportPathTips.Text = string.Format("(导出路径:{0})", exportPath);
            }
        }

        private void exportBtn_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(sourcePath))
            {
                Utils.ShowMsg("请拖入Excel文件或点选Excel文件");
                return;
            }
            if(string.IsNullOrEmpty(exportPath))
            {
                Utils.ShowMsg("请选择导出路径");
                return;
            }
            ExcelParserCore.Parse(sourcePath, exportTypeList[exportTypeIndex], exportPath);
        }
    }

}
