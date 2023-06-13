using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelParserGUI
{
    internal class HandlerBase
    {
        string excelPath;
        string excelName;
        string exportPath;
        StringBuilder sb;
        string lastSheet;
        int lastRow;

        internal HandlerBase()
        {
            lastRow = -1;
            lastSheet = string.Empty;
            sb = new StringBuilder();
        }

        internal void SetExcelPath(string path)
        {
            excelPath = path;
            var paths = excelPath.Split('\\');
            var lastName = paths[paths.Length - 1];
            var arr = lastName.Split('.');
            excelName = arr[0];
        }

        internal void SetExportPath(string path)
        {
            exportPath = path;
        }

        internal void Handle(ICell cell) 
        {
            bool newSheet = false;
            bool newRow = false;
            if (cell != null)
            {
                if (cell.Sheet.SheetName != lastSheet)
                {
                    newSheet = true;
                    if (!string.IsNullOrEmpty(lastSheet))
                    {
                        OnExport(cell.Sheet,sb.ToString());
                        sb.Clear();
                    }
                    lastSheet = cell.Sheet.SheetName;
                }
                if (cell.RowIndex != lastRow) {
                    if (lastRow != -1)
                    {
                        newRow = true;
                    }
                    lastRow = cell.RowIndex;
                }
            }
            sb.Append(OnHandle(cell, newRow, newSheet));
        }

        internal virtual string OnHandle(ICell cell, bool newRow, bool newSheet)
        {
            var data = cell == null ? "" : cell.ToString();
            if (newRow || newSheet) {
                return data + "\n";
            }
            return data;
        }

        internal virtual void OnExport(ISheet sheet, string result)
        {
            Utils.SaveFile(exportPath+"\\"+excelName+".txt", result);
        }
    }
}
