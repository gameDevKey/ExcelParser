using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.UserModel;

namespace ExcelParserCore
{
    public class ExcelParser
    {

    }

    private void readExcelData(object filePath)
    {
        try
        {
            if (!File.Exists(filePath.ToString()))
            {
                throw new Exception("文件不存在!");
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
                        if (row.GetCell(k) != null)
                        {

                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "发生错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
