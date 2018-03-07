using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace com.amtec.action
{

    /// 用C#实现汉字转化为拼音 

    /// 实现的原理就是先将汉字转化成为内码，然后通过内码和拼音的对照来查找。 

    public class PinYin
    {

       
        /// <summary>

        /// 获得第一个汉字的首字母（大写）；

        /// </summary>

        /// <param name="cnChar"></param>

        /// <returns></returns>

        public static string getSpell(string cnChar)
        {

            byte[] arrCN = Encoding.Default.GetBytes(cnChar);

            if (arrCN.Length > 1)
            {

                int area = (short)arrCN[0];

                int pos = (short)arrCN[1];

                int code = (area << 8) + pos;

                int[] areacode = { 45217, 45253, 45761, 46318, 46826, 47010, 47297, 47614, 48119, 48119, 49062, 49324, 49896, 50371, 50614, 50622, 50906, 51387, 51446, 52218, 52698, 52698, 52698, 52980, 53689, 54481 };

                for (int i = 0; i < 26; i++)
                {

                    int max = 55290;

                    if (i != 25) max = areacode[i + 1];

                    if (areacode[i] <= code && code < max)
                    {

                        return Encoding.Default.GetString(new byte[] { (byte)(65 + i) });

                    }

                }
                return "*";
            }

            else return cnChar;

        }

    }

}



