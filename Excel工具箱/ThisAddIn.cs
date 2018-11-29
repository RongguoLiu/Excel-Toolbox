using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Excel工具箱
{
    public partial class ThisAddIn
    {
        DigitToChnText digitToChnText = new DigitToChnText();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public int[] CellSelector()
        {
            Excel.Range range;
            int[] RangePosition = new int[2];
            try
            {
                range = (Application.InputBox(Prompt: "命名与哪个单元格相关？", Type: 8));
                range = range.Range["A1"];
                RangePosition[0] = range.Row;
                RangePosition[1] = range.Column;
                return RangePosition;
            }
            catch
            {
                return new int[] { 0, 0 };
            }
        }

        public string IndexInChinese(int i)
        {
            return digitToChnText.Converter(i.ToString());
        }

        /// 本程序用于将小写数字变成大写中文数字
        /// 算法设计：黄晶
        /// 程序制作：黄晶
        /// 时间：2004年8月12日
        class DigitToChnText
        {
            private readonly char[] chnText;
            private readonly char[] chnDigit;

            public DigitToChnText()
            {
                chnText = new char[] { '零', '一', '二', '三', '四', '五', '六', '七', '八', '九' };
                chnDigit = new char[] { '十', '百', '千', '万', '亿' };
            }

            public string Converter(string strDigit)
            {
                // 检查输入数字
                decimal dec;
                try
                {
                    dec = decimal.Parse(strDigit);
                }
                catch (FormatException)
                {
                    throw new Exception("输入数字的格式不正确。");
                }
                catch (Exception e)
                {
                    throw e;
                }

                if (dec <= -10000000000000000m || dec >= 10000000000000000m)
                {
                    throw new Exception("输入数字太大或太小，超出范围。");
                }

                StringBuilder strResult = new StringBuilder();

                // 提取符号部分
                // '+'在最前
                if ("+" == strDigit.Substring(0, 1))
                {
                    strDigit = strDigit.Substring(1);
                }
                // '-'在最前
                else if ("-" == strDigit.Substring(0, 1))
                {
                    strResult.Append('负');
                    strDigit = strDigit.Substring(1);
                }
                // '+'在最后
                else if ("+" == strDigit.Substring(strDigit.Length - 1, 1))
                {
                    strDigit = strDigit.Substring(0, strDigit.Length - 1);
                }
                // '-'在最后
                else if ("-" == strDigit.Substring(strDigit.Length - 1, 1))
                {
                    strResult.Append('负');
                    strDigit = strDigit.Substring(0, strDigit.Length - 1);
                }

                // 提取整数和小数部分
                int indexOfPoint;
                if (-1 == (indexOfPoint = strDigit.IndexOf('.'))) // 如果没有小数部分
                {
                    strResult.Append(ConvertIntegral(strDigit));
                }
                else // 有小数部分
                {
                    // 先转换整数部分
                    if (0 == indexOfPoint) // 如果“.”是第一个字符
                    {
                        strResult.Append('零');
                    }
                    else
                    {
                        strResult.Append(ConvertIntegral(strDigit.Substring(0, indexOfPoint)));
                    }

                    // 再转换小数部分
                    if (strDigit.Length - 1 != indexOfPoint) // 如果“.”不是最后一个字符
                    {
                        strResult.Append('点');
                        strResult.Append(ConvertFractional(strDigit.Substring(indexOfPoint + 1)));
                    }
                }

                return strResult.ToString();
            }

            // 转换整数部分
            protected string ConvertIntegral(string strIntegral)
            {
                // 去掉数字前面所有的'0'
                // 并把数字分割到字符数组中
                char[] integral = ((long.Parse(strIntegral)).ToString()).ToCharArray();

                // 变成中文数字并添加中文数位
                StringBuilder strInt = new StringBuilder();

                int i;
                int digit;
                digit = integral.Length - 1;

                // 处理最高位到十位的所有数字
                for (i = 0; i < integral.Length - 1; i++)
                {
                    strInt.Append(chnText[integral[i] - '0']);

                    if (0 == digit % 4)     // '万' 或 '亿'
                    {
                        if (4 == digit || 12 == digit)
                        {
                            strInt.Append(chnDigit[3]); // '万'
                        }
                        else if (8 == digit)
                        {
                            strInt.Append(chnDigit[4]); // '亿'
                        }
                    }
                    else         // '十'，'百'或'千'
                    {
                        strInt.Append(chnDigit[digit % 4 - 1]);
                    }

                    digit--;
                }

                // 如果个位数不是'0'
                // 或者个位数为‘0’但只有一位数
                // 则添加相应的中文数字
                if ('0' != integral[integral.Length - 1] || 1 == integral.Length)
                {
                    strInt.Append(chnText[integral[i] - '0']);
                }

                // 遍历整个字符串
                i = 0;
                while (i < strInt.Length)
                {
                    int j = i;

                    bool bDoSomething = false;

                    // 查找所有相连的“零X”结构
                    while (j < strInt.Length - 1 && "零" == strInt.ToString().Substring(j, 1))
                    {
                        string strTemp = strInt.ToString().Substring(j + 1, 1);

                        // 如果是“零万”或者“零亿”则停止查找
                        if ("万" == strTemp || "亿" == strTemp)
                        {
                            bDoSomething = true;
                            break;
                        }

                        j += 2;
                    }

                    if (j != i) // 如果找到“零X”结构，则全部删除
                    {
                        strInt = strInt.Remove(i, j - i);

                        // 除了在最尾处，或后面不是"零万"或"零亿"的情况下, 
                        // 其他处均补入一个“零”
                        if (i <= strInt.Length - 1 && !bDoSomething)
                        {
                            strInt = strInt.Insert(i, '零');
                            i++;
                        }
                    }

                    if (bDoSomething) // 如果找到"零万"或"零亿"结构
                    {
                        strInt = strInt.Remove(i, 1); // 去掉'零'
                        i++;
                        continue;
                    }

                    // 指针每次可移动2位
                    i += 2;
                }

                // 遇到“亿万”变成“亿零”或"亿"
                int index = strInt.ToString().IndexOf("亿万");
                if (-1 != index)
                {
                    if (strInt.Length - 2 != index &&  // 如果"亿万"不在最后
                     (index + 2 < strInt.Length && "零" != strInt.ToString().Substring(index + 2, 1))) // 并且其后没有"零"
                        strInt = strInt.Replace("亿万", "亿零", index, 2);
                    else
                        strInt = strInt.Replace("亿万", "亿", index, 2);
                }

                // 开头为“一十”改为“十”
                if (strInt.Length > 1 && "一十" == strInt.ToString().Substring(0, 2))
                {
                    strInt = strInt.Remove(0, 1);
                }

                return strInt.ToString();
            }

            // 转换小数部分
            protected string ConvertFractional(string strFractional)
            {
                char[] fractional = strFractional.ToCharArray();
                StringBuilder strFrac = new StringBuilder();

                // 变成中文数字
                int i;
                for (i = 0; i < fractional.Length; i++)
                {
                    strFrac.Append(chnText[fractional[i] - '0']);
                }

                return strFrac.ToString();
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
