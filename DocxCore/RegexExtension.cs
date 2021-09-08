using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Docx.Core
{
    public static class RegexExtension
    {
        /// <summary>
        /// 获取所有匹配的字符串的Match
        /// </summary>
        /// <param name="text"></param>
        /// <param name="regex"></param>
        /// <returns></returns>
        public static List<Match> GetMatches(this string text, Regex regex)
        {
            var matches = regex.Matches(text);
            return matches.ToList();
        }
        /// <summary>
        /// 获取字符串中第一个匹配正则的InnerText
        /// </summary>
        /// <param name="text"></param>
        /// <param name="regex"></param>
        /// <returns></returns>
        public static string FirstMatchOrDefault(this string text, Regex regex)
        {
            var match = regex.Match(text);
            if (!match.Success)
            {
                return string.Empty;
            }

            var innerText = match.Value;
            return innerText;
        }

        /// <summary>
        /// 获取转义后的正则表达式字符串
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetPattern(this string str)
        {
            var result = string.Empty;
            foreach (var s in str)
            {
                if (NeedEscape(s))
                {
                    result += @"\";
                }

                result += s.ToString();
            }

            return result;
        }

        /// <summary>
        /// 在正则表达式中是否需要转义
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool NeedEscape(this char s)
        {
            return s == '[' || s == ']' || s == '(' || s == ')' || s == '{' || s == '}' || s == '$' || s == '*' ||
                   s == '+' || s == '.' || s == '?' || s == 92 || s == '^' || s == '|';
        }

        /// <summary>
        /// 获取匹配列表字符串的所有字符串的Regex
        /// </summary>
        /// <param name="textList"></param>
        /// <returns></returns>
        public static Regex GetListRegex(this List<string> textList)
        {
            var pattern = string.Join("|", textList);
            return new Regex(pattern);
        }
        /// <summary>
        /// 获取匹配中间内容的Regex（非贪婪匹配）
        /// </summary>
        /// <param name="left">需要自行转义</param>
        /// <param name="right">需要自行转义</param>
        /// <param name="inner">需要自行转义</param>
        /// <returns></returns>
        public static Regex GetInnerTextRegex(string left, string right, string inner = "")
        {
            if (string.IsNullOrEmpty(inner))
            {
                inner = ".*?";
            }

            var pattern = "(?<=" + left + ")(" + inner + ")(?=" + right + ")";
            var result = new Regex(pattern);
            return result;
        }

        /// <summary>
        /// 多组值同时替换（简单用Replace是不行的，比如Value1=Key2，那么Key1就被替换为Value2了）
        /// </summary>
        /// <param name="text"></param>
        /// <param name="dictionary"></param>
        /// <returns></returns>
        public static string ReplaceByDictionary(this string text, Dictionary<string, string> dictionary)
        {
            var regex = dictionary.Keys.ToList().GetListRegex();
            var result = text.ReplaceByDictionary(dictionary, regex);
            return result;
        }

        private static string ReplaceByDictionary(this string text, Dictionary<string, string> dictionary, Regex regex)
        {
            var match = regex.Match(text);
            if (!match.Success)
            {
                return text;
            }

            var value = dictionary.FirstOrDefault(c => c.Key == match.Value).Value;
            var cut = match.Index + match.Value.Length;
            string result;
            if (string.IsNullOrEmpty(value))
            {
                result = text.Substring(0, cut);
            }
            else
            {
                result = text.Substring(0, match.Index) + value;
            }

            result += text.Substring(cut).ReplaceByDictionary(dictionary, regex);
            return result;
        }
    }
}
