using JIF.Common.Test.Entites;
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

namespace JIF.Common.Test
{
    [TestClass]
    public class NpoiExcelHelperTest
    {
        public readonly string ResultFile = @"C:\Users\Administrator\Desktop\npoiTestOutput\";

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

            excel.Write(res, 0, 0, 0);

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
            Assert.Fail("未进行测试");

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

            excel.Write(data, 0, 0, 0);

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

            /*
             * 注意 :
             * 使用 用户自定义实体类型写入Excel时,会自动根据 属性名称字母排序数据列
             */
            excel.Write(data, 0, 0, 0);

            excel.Export(string.Format("{0}Test_writeCustomerTypeListObject.xls", ResultFile));
        }

        [TestMethod]
        public void Test_wirteDynamicList()
        {
            NpoiExcelHelper excel = new NpoiExcelHelper();
            excel.Write(new[] { "第一列", "第二列", "三", "肆" }, 0, 0, 0);

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


            excel.Write(data, 0, 1, 0);

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
                c = d.Price,
            }).ToList();

            NpoiExcelHelper excel = new NpoiExcelHelper();
            excel.Write(data, 0, 0, 0);
            excel.Export(string.Format("{0}Test_writeAnonymousList.xls", ResultFile));

        }

        [TestMethod]
        public void Test_setStyle()
        {
            Assert.Fail("未进行测试");

            //List<dynamic> data = new List<dynamic>();

            //for (int i = 0; i < 10000; i++)
            //{
            //    dynamic o = new ExpandoObject();

            //    o.A = "Hello World";
            //    o.B = DateTime.Now.ToString("yyyy-MM-dd");
            //    o.C = 1.1m;
            //    o.D = 2.2d;

            //    data.Add(o);
            //}

            //NpoiExcelHelper excel = new NpoiExcelHelper();

            //excel.Write(data, 0, 0, 0);

            //ICellStyle footerCellstyle = excel.GetWorkBook().CreateCellStyle();
            //footerCellstyle.FillForegroundColor = HSSFColor.Red.Index;
            //footerCellstyle.FillPattern = FillPattern.SolidForeground;
            //excel.SetStyle(footerCellstyle, rowIndex: 5);

            //excel.Export(string.Format("{0}Test_setStyle.xls", ResultFile));

        }

        [TestMethod]
        public void Test_MultiTheadWriteList()
        {
            Assert.Fail("未进行测试");

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

        [TestMethod]
        public void 读取_返回List_Dynamic()
        {
            var file = @"F:\WorkDocument\Code.zen\Code.Zen.Test\assert\GCM Summer Sale_hotellist_Zeropartner.xlsx";
            var data = NpoiExcelHelper.Read(file, 0, 0, 0);

            var a = data.Select(d => new
            {
                hotelcode = d.C,
                hot = string.IsNullOrWhiteSpace(d.H) ? 0 : d.H,
                hotelbookingurl = d.Z,

            });

            var rs = JsonConvert.SerializeObject(a);

            Console.WriteLine(rs);
        }

        [TestMethod]
        public void Hilton_Summer_Olympics_Source_Test()
        {
            var file = @"E:\WorkDocument\Document\2016-05-13 夏季活动\包含酒店中文名,参加夏季活动酒店信息表格-0512.xlsx";
            var data = NpoiExcelHelper.Read(file, rowIndex: 1);

            var a = data.Select(d => new
            {
                hoteltype = d.AD,
                hotelname_ch = d.C,
                hotelcode = d.D,
                hotelcity = d.AE,
                hot = string.IsNullOrWhiteSpace(d.I) ? 0 : d.I,
                hotelbookingurl = d.AA,
            });

            var rs = JsonConvert.SerializeObject(a);

            Console.WriteLine(rs);
        }

        [TestMethod]
        public void BookLink_Mod()
        {
            var file = @"E:\WorkDocument\Document\2016-06-01 booklink 修改\6月1日需更新酒店信息表格-0523.xlsx";
            var data = NpoiExcelHelper.Read(file, rowIndex: 1).Select(d => new
            {
                HotelCode = d.D,
                BookLink = d.AB
            });

            StringBuilder sb = new StringBuilder();
            foreach (var item in data)
            {
                sb.AppendFormat("UNION SELECT '{0}','{1}'" + "\r\n", item.HotelCode, item.BookLink);
            }


            Console.WriteLine(sb.ToString());

            //Console.WriteLine(JsonConvert.SerializeObject(data));
        }

        [TestMethod]
        public void RFP_Hotel()
        {
            var file = @"E:\WorkDocument\Document\2016-06-03 希尔顿婚宴网站改版\RFP Email List - GCM Hotels New.xls";
            var data = NpoiExcelHelper.Read(file, rowIndex: 6).Select(d => new
            {
                HotelCode = d.D,
                Email = d.G
            });

            Console.WriteLine(JsonConvert.SerializeObject(data.Where(d => !string.IsNullOrWhiteSpace(d.Email)).OrderBy(d => d.Email)));
        }

        [TestMethod]
        public void Test_JSONConvert_Desc_Dynamic()
        {
            //string json = @"{ ""Name"":""陈宁"",""SEXx"":""男"" }";

            //var data = JsonConvert.DeserializeAnonymousType(json, new { Name = "", Sex = "", Age = 0 });

            //dynamic obj = new ExpandoObject();

            //obj.Name = "千里之外";
            //obj.Author = "JAY";
            //obj.Album = "2007世界巡回演唱会";
            //obj.FavoiCount = 167;

            //Console.WriteLine(JsonConvert.SerializeObject(obj));

            var json = @"[{ ""HotelCode"": ""SYXCICI"", ""Email"": ""Olina.Chen@hilton.com"" }, { ""HotelCode"": ""SYXDTDI"", ""Email"": ""Olina.Chen@hilton.com"" }, { ""HotelCode"": ""JHGXIDI"", ""Email"": ""airy.zhou@hilton.com"" }, { ""HotelCode"": ""BJSHITW"", ""Email"": ""Andrew.Moore@hilton.com"" }, { ""HotelCode"": ""HGHJIDI"", ""Email"": ""Andy.Gao@hilton.com"" }, { ""HotelCode"": ""SYXHIHI"", ""Email"": ""anna.cao@hilton.com"" }, { ""HotelCode"": ""XMNCICI"", ""Email"": ""anson.zengfl@conradhotels.com"" }, { ""HotelCode"": ""WUXTJDI"", ""Email"": ""audrey.zhang2@hilton.com"" }, { ""HotelCode"": ""CKGJBDI"", ""Email"": ""Barbara.Yu@hilton.com"" }, { ""HotelCode"": ""CSXZHHI"", ""Email"": ""Byrds.Yang@hilton.com"" }, { ""HotelCode"": ""HUZHADI"", ""Email"": ""Candy.Meng@hilton.com"" }, { ""HotelCode"": ""MFMCSCI"", ""Email"": ""carolina.cheung@sands.com.mo"" }, { ""HotelCode"": ""CZXWDHI"", ""Email"": ""chris.qian@hilton.com"" }, { ""HotelCode"": ""BJSWFHI"", ""Email"": ""Christina.Yang@hilton.com"" }, { ""HotelCode"": ""FOCLDDI"", ""Email"": ""Chuck.Fu@hilton.com"" }, { ""HotelCode"": ""TAOLXDI"", ""Email"": ""Coco.Zhu@hilton.com"" }, { ""HotelCode"": ""WUHRSHI"", ""Email"": ""Connie.Chao@hilton.com"" }, { ""HotelCode"": ""CANSCDI"", ""Email"": ""Demi.Li@hilton.com"" }, { ""HotelCode"": ""XMNWBDI"", ""Email"": ""denny.li@hilton.com"" }, { ""HotelCode"": ""HAKHCDI"", ""Email"": ""duke.du@hilton.com"" }, { ""HotelCode"": ""SYXHQDI"", ""Email"": ""eggie.lee@hilton.com"" }, { ""HotelCode"": ""DLUGTHI"", ""Email"": ""Elaine.Mo@hilton.com"" }, { ""HotelCode"": ""HKGHCCI"", ""Email"": ""Enoch.chiu@conradhotels.com"" }, { ""HotelCode"": ""SHAWAWA"", ""Email"": ""estella.xu@waldorfastoria.com"" }, { ""HotelCode"": ""SHAHITW"", ""Email"": ""Eva.fan@hilton.com"" }, { ""HotelCode"": ""DDGZDGI"", ""Email"": ""Eva.liang@hilton.com"" }, { ""HotelCode"": ""CANSRDI"", ""Email"": ""Frankie.liu@hilton.com"" }, { ""HotelCode"": ""CTUCCHI"", ""Email"": ""freda.liu@hilton.com"" }, { ""HotelCode"": ""SHASPDI"", ""Email"": ""Gavin.Dai@hilton.com"" }, { ""HotelCode"": ""HAKWEHI"", ""Email"": ""Gavin.hu@hilton.com"" }, { ""HotelCode"": ""NKGWUDI"", ""Email"": ""Grace.Zhu@hilton.com"" }, { ""HotelCode"": ""BJSCICI"", ""Email"": ""Haidy.Cheng@conradhotels.com"" }, { ""HotelCode"": ""HAKMEHI"", ""Email"": ""HAKME_CB@Hilton.com "" }, { ""HotelCode"": ""NGBNCDI"", ""Email"": ""Helen.He@hilton.com"" }, { ""HotelCode"": ""URCHHHI"", ""Email"": ""Jacky.He@hilton.com"" }, { ""HotelCode"": ""CANGUHI"", ""Email"": ""janica.shao@hilton.com"" }, { ""HotelCode"": ""SHASHHI"", ""Email"": ""Jenny.Qiu@hilton.com "" }, { ""HotelCode"": ""NKGJFHI"", ""Email"": ""Jenny.Xu@hilton.com"" }, { ""HotelCode"": ""SZXSBGI"", ""Email"": ""Karen.zhong@hilton.com"" }, { ""HotelCode"": ""HGHHEDI"", ""Email"": ""Kenny.Zhang@hilton.com"" }, { ""HotelCode"": ""SZVTVDI"", ""Email"": ""Lance.yan@hilton.com"" }, { ""HotelCode"": ""TAOGBHI"", ""Email"": ""Lily.Wang@hilton.com"" }, { ""HotelCode"": ""SZXSFHI"", ""Email"": ""lina.wang4@hilton.com"" }, { ""HotelCode"": ""HGHLRHI"", ""Email"": ""lion.mao@hilton.com"" }, { ""HotelCode"": ""LJGGIGI"", ""Email"": ""Lora.Luo@hilton.com"" }, { ""HotelCode"": ""CKGCWDI"", ""Email"": ""Lynn.Lingjuan@hilton.com"" }, { ""HotelCode"": ""TNAJHHI"", ""Email"": ""Maria.ma@hilton.com"" }, { ""HotelCode"": ""BJSWAWA"", ""Email"": ""Mark.Xu@waldorfastoria.com"" }, { ""HotelCode"": ""XIYHIHI"", ""Email"": ""Nancy.Wul@hilton.com"" }, { ""HotelCode"": ""HSNZHHI"", ""Email"": ""Nick.Zhao@hilton.com"" }, { ""HotelCode"": ""CGOZHHI"", ""Email"": ""Nicole.duan@hilton.com"" }, { ""HotelCode"": ""TSNECHI"", ""Email"": ""Pauline.Chu@hilton.com"" }, { ""HotelCode"": ""FUOCDHI"", ""Email"": ""Rachel.Zhu@hilton.com"" }, { ""HotelCode"": ""BJSDTDI"", ""Email"": ""Ray.Dailei@hilton.com"" }, { ""HotelCode"": ""GOQQG"", ""Email"": ""Rose.Fan@hilton.com"" }, { ""HotelCode"": ""NKGNRHI"", ""Email"": ""Sales.coordinator@hilton.com"" }, { ""HotelCode"": ""NKGYFHI"", ""Email"": ""salesc@hiltonfuxianlake.com"" }, { ""HotelCode"": ""SHEDTDI"", ""Email"": ""Seven.peng@hilton.com"" }, { ""HotelCode"": ""CANGTHI"", ""Email"": ""sita.chen@hilton.com"" }, { ""HotelCode"": ""WUHOVHI"", ""Email"": ""Stacy.Wang@hilton.com"" }, { ""HotelCode"": ""WUXXDDI"", ""Email"": ""sunshine.feng@hilton.com"" }, { ""HotelCode"": ""SJWZSHI"", ""Email"": ""Susan.Gao@hilton.com"" }, { ""HotelCode"": ""SZXSSHI"", ""Email"": ""Susie.cai@hilton.com"" }, { ""HotelCode"": ""BJSCAHI"", ""Email"": ""Susie.Sun2@hilton.com"" }, { ""HotelCode"": ""ZGNZDHI"", ""Email"": ""tiffany.mai@hilton.com"" }, { ""HotelCode"": ""HFEHIHI"", ""Email"": ""todd.fan@hilton.com"" }, { ""HotelCode"": ""HAKHAHI"", ""Email"": ""Vince.Cao@hilton.com"" }, { ""HotelCode"": ""AVAHSDT"", ""Email"": ""Vincent.Feng@hilton.com"" }, { ""HotelCode"": ""CKGHIHI"", ""Email"": ""Wendy.Lv@hilton.com"" }, { ""HotelCode"": ""TNAJCHI"", ""Email"": ""will.nie@hilton.com"" }, { ""HotelCode"": ""YNTDIDI"", ""Email"": ""Yongwei.wang@hilton.com"" }]";


            var a = JsonConvert.DeserializeObject<List<C1>>(json);

            var b = NpoiExcelHelper.Read(@"E:\WorkDocument\Document\2016-06-14 希尔顿婚宴网站- 酒店增加email\RFP Email List - GCM Hotels New 20160613.xls", rowIndex: 6).Select(d => new C1
            {
                HotelCode = d.D,
                Email = d.H
            });

            foreach (var item in a)
            {

                var c = b.FirstOrDefault(d => d.HotelCode.Trim() == item.HotelCode.Trim() && !string.IsNullOrWhiteSpace(d.Email));

                if (c == null) continue;

                item.Email += ";" + b.FirstOrDefault(d => d.HotelCode.Trim() == item.HotelCode.Trim()).Email;
            }

            Console.WriteLine(JsonConvert.SerializeObject(a));

        }


        public class C1
        {
            public string HotelCode { get; set; }

            public string Email { get; set; }
        }

    }
}