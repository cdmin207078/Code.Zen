using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JIF.Common.Extend
{
    public static class StringExtend
    {
        public static int ToAscii(this string character)
        {
            if (character.Length == 1)
            {
                ASCIIEncoding asciiEncoding = new ASCIIEncoding();

                return (int)asciiEncoding.GetBytes(character)[0];
            }
            else
                return -1;
        }
    }
}
