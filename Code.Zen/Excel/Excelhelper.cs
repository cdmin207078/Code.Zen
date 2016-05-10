using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.IO;
using System.Dynamic;

namespace JIF.Common.Excel
{
    public static class ExcelHelper
    {
        public static List<dynamic> Read(string file, int cellIndex = 0, int rowIndex = 0, int sheetIndex = 0)
        {
            IWorkbook workbook = null;

            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }

            if (workbook == null)
            {
                throw new Exception("没有读取到有效Excel数据");
            }

            var sheet = workbook.GetSheetAt(sheetIndex);
            var hChar = getExcelHeadChar(sheet.GetRow(rowIndex).LastCellNum);

            List<dynamic> result = new List<dynamic>();

            //遍历数据行
            for (int r = rowIndex; r <= sheet.LastRowNum; r++)
            {
                dynamic dyData = new ExpandoObject();
                var DicdyData = dyData as IDictionary<string, object>;

                IRow row = sheet.GetRow(r);

                //遍历一行的每一个单元格
                for (int c = cellIndex; c < row.LastCellNum; c++)
                {
                    ICell cel = row.GetCell(c);
                    if (cel == null)
                    {
                        DicdyData[hChar[c]] = null;
                    }
                    else
                    {
                        DicdyData[hChar[c]] = cel.ToString();
                    }
                }

                result.Add(dyData);
            }

            return result;
        }

        private static List<string> getExcelHeadChar(int length)
        {
            if (length < 1) return null;

            var result = new List<string>();

            for (int i = 1; i <= length; i++)
            {
                result.Add(Utils.ToNumberSystem26(i));
            }

            return result;
        }

        public static void Write<T>(List<T> data, int cellIndex = 0, int rowIndex = 0, int sheetIndex = 0)
        {

        }
    }
}
