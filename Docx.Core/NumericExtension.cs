using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Docx.Core
{
    /// <summary>
    /// 数值类型扩展
    /// </summary>
    public static class NumericExtension
    {
        private static readonly string[] hanzi = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
        //private static readonly List<char> hanziC = new List<char> { '零', '一', '二', '三', '四', '五', '六', '七', '八', '九' };
        // 进制单位 亿,兆,京,垓,姊,穰,沟,涧,正,载
        private static readonly string[] units = { "", "万", "亿", "兆", "京", "垓", "穰", "沟", "涧", "正", "载" };

        /// <summary>
        /// 转换为中文数值字符串
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToChinese(this int input)
        {
            return ToChinese((long)input);
        }
        /// <summary>
        /// 转换为中文数值字符串，中数(万进系统)
        /// </summary>
        /// <param name="input">数值</param>
        /// <returns>中文数值字符串</returns>
        public static string ToChinese(this long input)
        {
            string result;
            if (input == 0)
            {
                result = "零";
            }
            else if (input < 0)
            {
                result = "负" + ToChinese(-input);
            }
            else
            {
                result = ToChinese(input.ToString());
            }

            return result;
        }

        // 载正涧沟穰姊垓京兆亿万千百十 零一二三四五六七八九
        private static string ToChinese(string input)
        {
            // 目标 "9123400000560" -> "九兆一千二百三十四亿零五百六"
            var numStr = input;
            // "9123400000560" -> "0009123400000560"
            numStr = numStr.PadLeft((int)Math.Ceiling((double)numStr.Length / 4) * 4, '0');
            // 按四位分割成 {"0560","0000","1234","0009"}
            var numStrs = new List<string>();
            for (int i = numStr.Length; i > 0; i -= 4)
            {
                numStrs.Add(numStr.Substring(i - 4, 4));
            }
            var index = -1;
            var res = numStrs.Select(full =>  // 注：Select里面的代码不会立即执行，只有调用First和ToList这些方法的使用才会执行，但是每次调用里面的代码都会执行一遍
            {
                index++;
                return $"{hanzi[full[0] - 48]}千{hanzi[full[1] - 48]}百{hanzi[full[2] - 48]}十{hanzi[full[3] - 48]}{units[index]}";
            }).ToList().Reduce((left, right) =>
            {
                return right + left;
            });
            // "三千零百零十零" -> "三千零零十零"
            res = Regex.Replace(res, "零[载正涧沟穰姊垓京兆亿万千百十]", "零");
            // "三千零零十零" -> "三千零十零"
            res = Regex.Replace(res, "零+", "零");
            // 去掉末尾的零 "250" -> "二百五十零"
            res = Regex.Replace(res, "(^零)|(零$)", "");
            // 可选 去掉末尾连续单位的最后一个单位 "二百五十"->"二百五","三万六千"->"三万六"
            res = Regex.Replace(res, "([载正涧沟穰姊垓京兆亿万千百])([一二三四五六七八九])([正涧沟穰姊垓京兆亿万千百十])$", "$1$2");
            // 可选 去掉开头“一十”的“一” "一十九"->"十九" "一十九万"->"十九万"
            res = Regex.Replace(res, "^一十", "十");
            return res;
        }

        /// <summary>
        /// [x1, x2, x3, x4].Reduce(f) = f(f(f(x1, x2), x3), x4).
        /// </summary>
        /// <typeparam name="T">The element type.</typeparam>
        /// <param name="source">Source IEnumerable.</param>
        /// <param name="merge">Merge expression.</param>
        /// <returns>return default if source is null or not exist any element.</returns>
        public static T Reduce<T>(this IEnumerable<T> source, Func<T, T, T> merge)
        {
            if (source == null || !source.Any()) return default;
            var result = source.First();
            foreach (var item in source.Skip(1))
            {
                result = merge(result, item);
            }
            return result;
        }
    }
}
