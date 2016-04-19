using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace Code.Zen
{
    public interface IExcelHelper
    {
        IWorkbook GetWorkBook();

        void CreateSheet(string sheetName);
        void CreateRow(int sheetIndex, int rowIndex);

        void Write(IList<dynamic> source, int sheetIndex, int rowIndex, int CellIndex);
        void Write<T>(T[] source, int sheetIndex, int rowIndex, int CellIndex);
        void Write<T>(T[,] source, int sheetIndex, int rowIndex, int CellIndex);
        void Write<T>(IList<T> source, int sheetIndex, int rowIndex, int CellIndex);

        void Write<T>(T value, int rowIndex, int CellIndex);
        void Write<T>(T value, int sheetIndex, int rowIndex, int CellIndex);

        void Export(string filePath);

        T Read<T>(string filePath, int sheetIndex, int rowIndex, int CellIndex);
        IList<T> ReadList<T>(string filePath, int sheetIndex, int rowIndex, int CellIndex, int endRowIndex, int endCellIndex);

        List<dynamic> ReadAsDynamicList(string filePath);

        void SetStyle(ICellStyle style, int sheetIndex, int rowIndex);
        void SetStyle(ICellStyle style, int sheetIndex, int rowIndex, int CellIndex);
    }
}
