using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.IO;
using System.Dynamic;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace JIF.Common.Excel
{
    public class NpoiExcelHelper
    {
        private IWorkbook _workbook;

        public NpoiExcelHelper(bool IsExcel2007 = false)
        {
            if (IsExcel2007)
            {
                _workbook = new XSSFWorkbook();
            }
            else
            {
                _workbook = new HSSFWorkbook();
            }


            CreateSheet("Sheet1");
        }

        #region Private

        ISheet Sheet(int sheetIndex)
        {
            return _workbook.GetSheetAt(sheetIndex);
        }

        IRow Row(int rowIndex, int sheetIndex)
        {
            return Sheet(sheetIndex).GetRow(rowIndex);
        }

        ICell Cell(int cellIndex, int rowIndex, int sheetIndex)
        {
            return Row(sheetIndex: sheetIndex, rowIndex: rowIndex).GetCell(cellIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        #endregion

        public void CreateSheet(string sheetName)
        {
            _workbook.CreateSheet(sheetName);
        }

        public void CreateRow(int sheetIndex, int rowIndex)
        {
            _workbook.GetSheetAt(sheetIndex).CreateRow(rowIndex);
        }

        public void Write<T>(T source, int sheetIndex, int rowIndex, int cellIndex)
        {
            if (Row(rowIndex, sheetIndex) == null)
                CreateRow(sheetIndex: sheetIndex, rowIndex: rowIndex);

            var currentCell = Cell(cellIndex, rowIndex, sheetIndex);

            if (source != null)
            {
                var tp = source.GetType();

                if (tp == typeof(decimal) || tp == typeof(double) || tp == typeof(float) || tp == typeof(int))
                {
                    currentCell.SetCellValue(Convert.ToDouble(source));
                }
                //else if (tp == typeof(DateTime))
                //{
                //    //IDataFormat format = _workbook.CreateDataFormat();
                //    //currentCell.CellStyle.DataFormat = format.GetFormat("yyyy-MM-dd HH:mm:ss");

                //    currentCell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("yyyy-MM-dd HH:mm:ss");
                //    currentCell.SetCellValue(Convert.ToDateTime(value));
                //}
                else
                {
                    currentCell.SetCellValue(source.ToString());
                }
            }
            else
            {
                currentCell.SetCellValue("");
            }
        }

        public void Write<T>(T[] source, int sheetIndex, int rowIndex, int cellIndex)
        {
            if (source == null || source.Length == 0)
                return;

            for (int i = 0; i < source.Length; i++)
            {
                Write(source[i], sheetIndex, rowIndex, cellIndex + i);
            }
        }

        public void Write<T>(T[,] source, int sheetIndex, int rowIndex, int cellIndex)
        {
            if (source == null || source.Length == 0)
                return;

            var row = source.GetLength(0);
            var col = source.GetLength(1);

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    Write(source[i, j], sheetIndex, rowIndex + i, cellIndex + j);
                }
            }
        }

        public void Write<T>(List<T> source, int sheetIndex, int rowIndex, int cellIndex)
        {
            if (source == null || source.Count() == 0)
                return;

            var type = typeof(T);

            if (type == typeof(ValueType) || type == typeof(string))
            {
                Write(source.ToArray(), sheetIndex, rowIndex, cellIndex);
            }
            else
            {
                var props = typeof(T).GetProperties();
                for (int i = 0; i < source.Count(); i++)
                {
                    for (int j = 0; j < props.Length; j++)
                    {
                        Write(props[j].GetValue(source[i], null), sheetIndex, rowIndex + i, cellIndex + j);
                    }
                }
            }
        }

        public void Write(List<dynamic> source, int sheetIndex, int rowIndex, int cellIndex)
        {
            if (source == null || source.Count() == 0)
                return;

            for (int i = 0; i < source.Count(); i++)
            {
                int col = 0;
                foreach (var initem in source[i])
                {
                    Write(initem.Value, sheetIndex, rowIndex + i, cellIndex + col);
                    col++;
                }
            }
        }

        public void Export(string filePath)
        {
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    _workbook.Write(fs);
                }
            }
            _workbook = null;
        }

        public static List<dynamic> Read(string file, int sheetIndex, int rowIndex, int cellIndex)
        {
            IWorkbook workbook = null;

            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(fs);
            }

            if (workbook == null)
            {
                throw new Exception(" 没有读取到有效Excel数据 - NpoiExcelHelper.Read ");
            }

            var sheet = workbook.GetSheetAt(sheetIndex);

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
                        DicdyData[Utils.ToNumberSystem26(c)] = null;
                    }
                    else
                    {
                        DicdyData[Utils.ToNumberSystem26(c)] = cel.ToString();
                    }
                }

                result.Add(dyData);
            }

            return result;
        }
    }
}
