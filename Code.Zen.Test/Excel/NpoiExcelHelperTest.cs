using JIF.Common.Test.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JIF.Common.Excel.Test
{
    [TestClass]
    public class NpoiExcelHelperTest
    {
        public readonly string ResultFile = @"C:\Users\Administrator\Desktop\npoiTestOutput\";

        [TestMethod]
        public void 写入_10000x100()
        {
            IExcelHelper excel = new NpoiExcelHelper();
            excel.CreateSheet("1-1");

            for (int i = 0; i < 10000; i++)
            {
                excel.CreateRow(0, i);
                for (int j = 0; j < 100; j++)
                {
                    excel.Write(i * j, 0, i, j);
                }
            }

            excel.Export(string.Format("{0}写入数据_10000x100.xls", ResultFile));
        }

        [TestMethod]
        public void Test_writeStringArray()
        {
            string[,] res = new string[,] {
                {"A","B","C","D","E","F"},
                {"A1","B2","C3","D4","E5","F6"},
                {"A11","B22","C33","D44","E55","F66"},
            };

            NpoiExcelHelper excel = new NpoiExcelHelper();
            excel.CreateSheet("1-1");

            excel.Write(res, 0, 3, 3);

            excel.Export(string.Format("{0}Test_writeStringArray.xls", ResultFile));
        }

        [TestMethod]
        public void Test_writeNumberArray()
        {
            double[,] res = new double[,] {
                {1,2,3,4,5,6,7,8},
                {1.1,2.2,3.3,4.4,5.5,6.6,7.7,8.8},
                {1.11,2.22,3.33,4.44,5.55,6.66,7.77,8.88},
            };

            NpoiExcelHelper excel = new NpoiExcelHelper();
            excel.CreateSheet("1-1");

            excel.Write(res, 0, 3, 3);

            excel.Export(string.Format("{0}Test_writeNumberArray.xls", ResultFile));
        }

        [TestMethod]
        public void Test_writeDataTable()
        {
            //DataTable dt = new DataTable();
            //dt.Columns.Add("string", typeof(string));
            //dt.Columns.Add("int", typeof(int));
            //dt.Columns.Add("decimal", typeof(decimal));
            //dt.Columns.Add("double", typeof(double));
            //dt.Columns.Add("datetime", typeof(DateTime));


            //string[] colTitles = new List<string>() { "标题", "int 数值", "decimal 数值", "double 数值", "datetime 数值" }.ToArray();

            //for (int i = 0; i < 30000; i++)
            //{
            //    DataRow dr = dt.NewRow();
            //    dr["string"] = "NPOI向Excel文件中插入数值时，可能会出现数字当作文本的情况（即左上角有个绿色三角），这样单元格的值就无法参与运算。这是因为在SetCellValue设置单元格值的时候使用了字符串进行赋值，默认被转换成了字符型。如果需要纯数字型的，请向SetCellValue中设置数字型变量。sheet.GetRow(2).CreateCell(2).SetCellValue(123);";
            //    dr["int"] = i;
            //    dr["decimal"] = Convert.ToDecimal(i) / 2;
            //    dr["double"] = Convert.ToDouble(i) / 2;
            //    dr["datetime"] = DateTime.Now;
            //    dt.Rows.Add(dr);
            //}

            //NpoiExcelHelper excel = new NpoiExcelHelper();

            //excel.Write(colTitles);
            //excel.Write(dt, rowIndex: 1);

            //excel.Export(string.Format("{0}Test_writeDataTable.xls", ResultFile));
        }

        [TestMethod]
        public void Test_writeListValueType()
        {
            List<string> data = new List<string>() { "A", "B", "C" };
            //List<List<string>> data = new List<List<string>>
            //{
            //    new List<string> {"A"},
            //    new List<string> {"B","C"},
            //    new List<string> {"D","E","F"}
            //};

            NpoiExcelHelper excel = new NpoiExcelHelper();


            excel.Write(data);

            excel.Export(string.Format("{0}Test_writeListValueType.xls", ResultFile));
        }

        [TestMethod]
        public void Test_writeCustomerTypeListObject()
        {
            var data = new List<Product>();

            for (int i = 0; i < 10000; i++)
            {
                data.Add(new Product
                {
                    SysNo = i,
                    ProductId = "编号:" + i,
                    Price = Convert.ToDecimal(i) / 3,
                    CreateTime = DateTime.Now
                });
            }

            NpoiExcelHelper excel = new NpoiExcelHelper();

            excel.Write(data);

            excel.Export(string.Format("{0}Test_writeCustomerTypeListObject.xls", ResultFile));
        }

        [TestMethod]
        public void Test_wirteDynamicList()
        {
            var data = new List<dynamic>();

            for (int i = 0; i < 10000; i++)
            {
                dynamic o = new ExpandoObject();

                o.A = "Hello World";
                o.B = DateTime.Now.ToString("yyyy-MM-dd");
                o.C = 1.1m;
                o.D = 2.2d;

                data.Add(o);
            }

            NpoiExcelHelper excel = new NpoiExcelHelper();

            excel.Write(data);

            excel.Export(string.Format("{0}Test_wirteDynamicList.xls", ResultFile));

        }

        [TestMethod]
        public void Test_writeAnonymousList()
        {
            var res = new List<Product>();

            for (int i = 0; i < 10000; i++)
            {
                res.Add(new Product
                {
                    SysNo = i,
                    ProductId = "编号:" + i,
                    Price = Convert.ToDecimal(i) / 3,
                    CreateTime = DateTime.Now
                });
            }

            var data = res.Select(d => new
            {
                B = d.ProductId,
                A = d.SysNo,

                c = d.Price
            }).ToList();

            NpoiExcelHelper excel = new NpoiExcelHelper();
            excel.Write(data);
            excel.Export(string.Format("{0}Test_writeAnonymousList.xls", ResultFile));

        }

        [TestMethod]
        public void Test_setStyle()
        {
            List<dynamic> data = new List<dynamic>();

            for (int i = 0; i < 10000; i++)
            {
                dynamic o = new ExpandoObject();

                o.A = "Hello World";
                o.B = DateTime.Now.ToString("yyyy-MM-dd");
                o.C = 1.1m;
                o.D = 2.2d;

                data.Add(o);
            }

            NpoiExcelHelper excel = new NpoiExcelHelper();

            excel.Write(data);

            ICellStyle footerCellstyle = excel.GetWorkBook().CreateCellStyle();
            footerCellstyle.FillForegroundColor = HSSFColor.Red.Index;
            footerCellstyle.FillPattern = FillPattern.SolidForeground;
            excel.SetStyle(footerCellstyle, rowIndex: 5);

            excel.Export(string.Format("{0}Test_setStyle.xls", ResultFile));

        }

        [TestMethod]
        public void Test_MultiTheadWriteList()
        {
            //Parallel.For(0, 1000, (i) =>
            //{
            //    var data = new List<Product>();

            //    for (int j = 0; j < 10000; j++)
            //    {
            //        data.Add(new Product
            //        {
            //            SysNo = j,
            //            ProductId = "编号:" + j,
            //            Price = Convert.ToDecimal(j) / 3,
            //            CreateTime = DateTime.Now
            //        });
            //    }
            //    try
            //    {
            //        NpoiExcelHelper excel = new NpoiExcelHelper();

            //        excel.Write(data);

            //        excel.Export(string.Format("{0}Test_writeCustomerTypeListObject" + i + "_.xls", ResultFile));
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex.Message);
            //        //throw;
            //    }
            //});

        }




        #region IExcel & NExcel Test 

        [TestMethod]
        public void Test_GetExcelColHeadChar()
        {
            var a = Utils.ToNumberSystem26(0);
            var b = Utils.ToNumberSystem26(10);
            var c = Utils.ToNumberSystem26(30);
            var d = Utils.ToNumberSystem26(1000);
        }


        [TestMethod]
        public void Test_NExcel_ReadExcel()
        {
            var file = @"E:\WorkDocument\Document\2016-05-13 夏季活动\GCM Summer Sale_hotellist_Zeropartner.xlsx";
            var data = ExcelHelper.Read(file, rowIndex: 1);

            var a = data.Select(d => new
            {
                hotelcode = d.C,
                hot = string.IsNullOrWhiteSpace(d.H) ? 0 : d.H,
                hotelbookingurl = d.Z
            });

            var rs = JsonConvert.SerializeObject(a);

            Console.WriteLine(rs);
        }

        #endregion
    }
}
