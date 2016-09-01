using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace JIF.Common.Test
{
    [TestClass]
    public class RandomHelperTest
    {


        [TestMethod]
        public void Test_GuidHashCode()
        {
            for (int i = 0; i < 100; i++)
            {
                var guid = Guid.NewGuid();
                Console.WriteLine(string.Format("guid: {0}, hcode: {1}", guid, guid.GetHashCode()));
            }
        }

        [TestMethod]
        public void Test_Generate()
        {
            for (int i = 0; i < 100; i++)
            {
                Console.WriteLine(string.Format("重复个数: {0}", 10000 - RandomHelper.Gen(RandomHelper.Format.NumChar, 4, 10000).Distinct().Count()));
            }
        }

        [TestMethod]
        public void Test_Generate_NumChar()
        {
            var data = RandomHelper.Gen(RandomHelper.Format.NumCharL, 4, 10000);

            foreach (var d in data)
            {
                Console.WriteLine(d);
            }

        }
    }
}
