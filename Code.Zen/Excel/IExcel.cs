using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JIF.Common.Excel
{
    interface IExcel
    {
        void Write<T>(List<T> data, int cellIndex = 0, int rowIndex = 0, int sheetIndex = 0);

        List<dynamic> Read(string filename, int cellIndex = 0, int rowIndex = 0, int sheetIndex = 0);
    }
}
