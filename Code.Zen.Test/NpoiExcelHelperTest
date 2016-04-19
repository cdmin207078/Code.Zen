using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Threading.Tasks;
using ISample.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.Linq;
using FluentEmail;

namespace Code.Zen.Test
{
    [TestClass]
    public class NpoiExcelHelperTest
    {
        public readonly string ResultFile = @"C:\Users\pc\Desktop\npoiTestOutput\";

        [TestMethod]
        public void Test_writeCellValue()
        {
            NpoiExcelHelper excel = new NpoiExcelHelper();
            excel.CreateSheet("1-1");

            for (int i = 0; i < 10000; i++)
            {
                excel.CreateRow(0, i);
                for (int j = 0; j < 100; j++)
                {
                    excel.Write(i * j, 0, i, j);
                }
            }

            excel.Export(string.Format("{0}Test_writeCellValue.xls", ResultFile));
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

        public class Product
        {
            public DateTime CreateTime { get; set; }
            public int SysNo { get; set; }
            public string ProductId { get; set; }
            public decimal Price { get; set; }
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
            Parallel.For(0, 1000, (i) =>
            {
                var data = new List<Product>();

                for (int j = 0; j < 10000; j++)
                {
                    data.Add(new Product
                    {
                        SysNo = j,
                        ProductId = "编号:" + j,
                        Price = Convert.ToDecimal(j) / 3,
                        CreateTime = DateTime.Now
                    });
                }
                try
                {
                    NpoiExcelHelper excel = new NpoiExcelHelper();

                    excel.Write(data);

                    excel.Export(string.Format("{0}Test_writeCustomerTypeListObject" + i + "_.xls", ResultFile));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    //throw;
                }
            });

        }

        [TestMethod]
        public void Test_ReadAndWriteExcelFile()
        {
            string fileName = string.Format("{0}BillTemplet.xls", ResultFile);
            NpoiExcelHelper excel = new NpoiExcelHelper(fileName);

            int tempDataStart = 21;
            int tempSumStart = 26;
            int OriginTotalRow = excel.Sheet(0).LastRowNum;

            #region Fill Table Title
            var st = excel.Sheet(0);

            var r = new Random(1);
            ICellStyle dataStyle = excel.Row(25).RowStyle;

            for (int i = 0; i < 2; i++)
            {

                int fillStart = excel.Sheet(0).PhysicalNumberOfRows;

                st.CopyRow(tempDataStart, fillStart);           //--title 1
                st.CopyRow(tempDataStart + 1, fillStart + 1);   //--title 2
                st.CopyRow(tempDataStart + 2, fillStart + 2);   //--标题栏

                int datacount = r.Next(100);
                //int datacount = 100;

                for (int j = 0; j < datacount; j++)
                {
                    int FillrowIndex = fillStart + j + 3;


                    st.CopyRow(tempDataStart + 3, FillrowIndex);   //--数据
                    //excel.CreateRow(rowIndex: FillrowIndex);

                    var row = excel.Row(FillrowIndex);

                    excel.Write("108056_" + j, FillrowIndex, 0);
                    excel.Write("那就对了，说明你开始上手了。", FillrowIndex, 1);
                    excel.Write("个", FillrowIndex, 9);

                }

                st.CopyRow(tempDataStart + 4, fillStart + datacount + 3);   //--小计

                st.CreateRow(fillStart + datacount + 4);
                st.CreateRow(fillStart + datacount + 5);
                //st.ShiftRows(fillStart, st.LastRowNum, 2, true, true);

            }

            #endregion


            st.SetActiveCellRange(tempDataStart, OriginTotalRow, 0, 12);

            //将汇总数据模板下移
            st.ShiftRows(tempSumStart, OriginTotalRow, st.LastRowNum - tempSumStart - 1, true, true);

            //将添加的数据区域 + 汇总区域 上移
            st.ShiftRows(OriginTotalRow, st.LastRowNum, 0 - (OriginTotalRow - tempDataStart + 1), true, true);

            excel.Export(string.Format("{0}BillTemplet_Res.xls", ResultFile));
        }

        [TestMethod]
        public void Test_Read_Dynamic_List()
        {
            string fileName = string.Format("{0}EmailNotice.xlsx", ResultFile);
            NpoiExcelHelper excel = new NpoiExcelHelper();

            List<dynamic> data = excel.ReadAsDynamicList(fileName);


            string toEmail = "cdmin207078@foxmail.com";
            string fromEmail = "johno@test.com";
            string subject = "Npoi Excel数据读取导入+FluentEmail发送邮件";

            string template = string.Format("收件人:@Model.收件人,邮件标题:@Model.邮件标题,内容模板:@Model.内容模板,数据区域:@Model.数据区域");
            //var template = string.Format("编号:@Model.编号,数量:@Model.数量,金额:@Model.金额");

            foreach (var item in data)
            {
                var email = Email
                .From(fromEmail)
                .To(toEmail)
                .Subject(subject)
                .UsingTemplate(template, item)
                .Send();

                Console.WriteLine(email.Message.Body);
            }
        }
    }
}
