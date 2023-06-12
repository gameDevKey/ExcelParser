using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.UserModel;

namespace ExcelParserGUI
{
    public class ExcelParserCore
    {
        public static void Parse(string filePath, string parseType)
        {
            var tpe = Config.Handlers[parseType];
            if (tpe == null)
            {
                Utils.ShowErrorMsg("找不到解析器:"+ parseType);
                return;
            }
            var cls = System.Activator.CreateInstance(tpe) as HandlerBase;
            cls.SetExcelPath(filePath);
            handleExcelData(filePath, cls.Handle);
        }

        private static void handleExcelData(string filePath, Action<ICell> handler)
        {
            try
            {
                if (!File.Exists(filePath.ToString()))
                {
                    Utils.ShowErrorMsg("文件不存在!");
                    return;
                }

                FileStream fsRead = new FileStream(filePath.ToString(), FileMode.Open);
                IWorkbook workBook = new HSSFWorkbook(fsRead);
                int sheetNum = workBook.NumberOfSheets;
                for (int i = 0; i < sheetNum; i++)
                {
                    ISheet sheet = workBook.GetSheetAt(i);
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum;    //一行最后一个cell的编号 即总的列数
                    for (int j = sheet.FirstRowNum; j <= sheet.LastRowNum; j++)
                    {
                        IRow row = sheet.GetRow(j);
                        if (row == null) continue;

                        for (int k = row.FirstCellNum; k < cellCount; ++k)
                        {
                            handler.Invoke(row.GetCell(k));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.ShowErrorMsg(ex.ToString());
            }
        }
    }

}
