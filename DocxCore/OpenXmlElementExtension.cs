using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Docx.Core
{
    public static class OpenXmlElementExtension
    {
        /// <summary>
        /// Docx的全局的命名空间，每个元素需要有形如特性xmlns:w=NameSpace
        /// </summary>
        public static readonly string NameSpace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        /// <summary>
        /// 给xml加上命名空间，仅用于从OuterXml中复制出来的xml元素，在ToOpenXmlElement之前调用
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static string AddNameSpace(this string xml)
        {
            xml = xml.Trim();
            var regex = RegexExtension.GetInnerTextRegex("<", ":");
            var ns = xml.FirstMatchOrDefault(regex);
            var spaceRegex = new Regex(" ");
            var replace = spaceRegex.Replace(xml, $" xmlns:{ns}=\"{NameSpace}\" ", 1);
            return replace;
        }

        /// <summary>
        /// Xml转OpenXmlElement
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static OpenXmlElement ToOpenXmlElement(this string xml)
        {
            using StreamWriter sw = new StreamWriter(new MemoryStream());
            sw.Write(xml);
            sw.Flush();
            sw.BaseStream.Seek(0, SeekOrigin.Begin);

            OpenXmlReader re = OpenXmlReader.Create(sw.BaseStream);

            re.Read();
            OpenXmlElement openXmlElement = re.LoadCurrentElement();
            re.Close();
            return openXmlElement;
        }
        /// <summary>
        /// XElement转OpenXmlElement
        /// </summary>
        /// <param name="xlElement"></param>
        /// <returns></returns>
        public static OpenXmlElement ToOpenXmlElement(this XElement xlElement)
        {
            return xlElement.ToString().ToOpenXmlElement();
        }
        /// <summary>
        /// xml转XElement
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static XElement ToXElement(this string xml)
        {
            var result = XElement.Parse(xml);
            return result;
        }
        /// <summary>
        /// OpenXmlElement转XElement
        /// </summary>
        /// <param name="openXmlElement"></param>
        /// <returns></returns>
        public static XElement ToXElement(this OpenXmlElement openXmlElement)
        {
            return openXmlElement.OuterXml.ToXElement();
        }
        /// <summary>
        /// 复制
        /// </summary>
        /// <param name="openXmlElement"></param>
        /// <returns></returns>
        public static OpenXmlElement GetClone(this OpenXmlElement openXmlElement)
        {
            return (OpenXmlElement)openXmlElement.Clone();
        }

        /// <summary>
        /// 节点替换，并返回父级节点
        /// </summary>
        /// <param name="openXmlElement"></param>
        /// <param name="replace"></param>
        /// <returns></returns>
        public static OpenXmlElement Replace<T>(this T openXmlElement, T replace) where T : OpenXmlElement
        {
            return Replace(openXmlElement, new List<OpenXmlElement> { replace });
        }
        /// <summary>
        /// 节点替换，并返回父级节点
        /// </summary>
        /// <param name="openXmlElement"></param>
        /// <param name="replaceList"></param>
        /// <returns></returns>
        public static OpenXmlElement Replace<T>(this T openXmlElement, List<T> replaceList) where T : OpenXmlElement
        {
            var last = openXmlElement;
            var parent = openXmlElement.Parent;
            foreach (var replace in replaceList)
            {
                if (replace == null)
                {
                    continue;
                }

                parent.InsertAfter(replace, last);
                last = replace;
            }

            parent.RemoveChild(openXmlElement);
            return parent;
        }

        /// <summary>
        /// 把一个OpenXmlElement里面的文本设置为指定值
        /// </summary>
        /// <param name="element"></param>
        /// <param name="innerText"></param>
        /// <returns>入参本身</returns>
        public static OpenXmlElement SetInnerText(this OpenXmlElement element, string innerText)
        {
            //数组的内容形如XX,X ,{,XXX,}, XXX，不能保证断句后{}里的内容一定是连续完整
            var texts = element.Descendants<Text>().ToList();
            var isFirst = true;
            foreach (var text in texts)
            {
                if (isFirst)
                {
                    //把内容给第一个，其余设置空
                    text.Text = innerText;
                    isFirst = false;
                }
                else
                {
                    text.Text = string.Empty;
                }
            }

            return element;
        }
        /// <summary>
        /// 设置最里面的节点的xml，非独立Element无效
        /// </summary>
        /// <param name="element"></param>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static OpenXmlElement SetInnermostXml(this OpenXmlElement element, string xml)
        {
            var innermost = element.GetInnermost();
            if (innermost == null)
            {
                return element;
            }

            var newInnermost = xml.ToOpenXmlElement();
            innermost.Replace(newInnermost);
            return element;
        }

        /// <summary>
        /// 获取最里面的OpenXmlElement，非独立Element无效，返回null
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static OpenXmlElement GetInnermost(this OpenXmlElement element)
        {
            if (!element.HasChildren)
            {
                return element;
            }

            var result = element.ChildElements.Count(c => !c.IsRpr()) == 1;
            if (!result)
            {
                //throw new Exception("Not A Independent OpenXmlElement");
                return null;
            }

            var child = element.ChildElements.First(c => !c.IsRpr());
            return child.GetInnermost();
        }

        /// <summary>
        /// 获取OpenXmlElement里面所有指定的独立的OpenXmlElement
        /// </summary>
        /// <param name="element"></param>
        /// <param name="judgeFunc"></param>
        /// <returns></returns>
        public static List<OpenXmlElement> GetIndependents(this OpenXmlElement element, Func<OpenXmlElement, bool> judgeFunc = null)
        {
            var result = new List<OpenXmlElement>();
            var res = element.IsIndependent(judgeFunc);
            if (res == IndependentResult.JudgeTrue)
            {
                result.Add(element);
                return result;
            }

            if (res == IndependentResult.JudgeFalse)
            {
                return result;
            }

            foreach (var childElement in element.ChildElements)
            {
                var temp = childElement.IsIndependent(judgeFunc);
                if (temp == IndependentResult.JudgeTrue)
                {
                    result.Add(childElement);
                }
                else if (temp == IndependentResult.JudgeFalse)
                {
                    //没必要再进去递归
                }
                else
                {
                    result.AddRange(childElement.GetIndependents(judgeFunc));
                }
            }

            return result;
        }

        /// <summary>
        /// 判断一个OpenXmlElement是否是独立的，即往下的子OpenXmlElement都只有一个（w:rPr的不算，这个是样式）
        /// </summary>
        /// <param name="element">待判断的Element</param>
        /// <param name="judgeFunc">对最里面一层的判断，不判断则返回true</param>
        /// <returns></returns>
        public static IndependentResult IsIndependent(this OpenXmlElement element, Func<OpenXmlElement, bool> judgeFunc = null)
        {
            if (!element.HasChildren)
            {
                if (judgeFunc == null || judgeFunc(element))
                {
                    return IndependentResult.JudgeTrue;
                }

                return IndependentResult.JudgeFalse;
            }

            var result = element.ChildElements.Count(c => !c.IsRpr()) == 1;
            if (!result)
            {
                return IndependentResult.IsNotIndependent;
            }

            var child = element.ChildElements.First(c => !c.IsRpr());
            return child.IsIndependent(judgeFunc);
        }

        /// <summary>
        /// 是否独立OpenXmlElement的结果
        /// </summary>
        public enum IndependentResult
        {
            /// <summary>
            /// 是独立节点，但判断为false
            /// </summary>
            JudgeFalse = 0,
            /// <summary>
            /// 是独立节点，且判断为true
            /// </summary>
            JudgeTrue = 1,
            /// <summary>
            /// 不是独立节点
            /// </summary>
            IsNotIndependent = 2
        }
        /// <summary>
        /// 是否是表示样式的节点
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static bool IsRpr(this OpenXmlElement element)
        {
            return element.Prefix == "w" && element.LocalName == "rPr";
        }

        /// <summary>
        /// 将文字替换为xml（特殊字符才这么替换）
        /// </summary>
        /// <param name="element"></param>
        /// <param name="text">文字</param>
        /// <param name="xml">xml</param>
        /// <param name="indexes">第几个text替换，从0开始计算，null则全部替换，空数组不会替换</param>
        /// <returns></returns>
        public static OpenXmlElement ReplaceInnerTextToXml(this OpenXmlElement element, string text, string xml,
            List<int> indexes = null)
        {
            if (indexes != null)
            {
                indexes = indexes.Where(c => c >= 0).ToList();
                if (!indexes.Any())
                {
                    return element;
                }
            }

            var openXmlElements = element.GetIndependents(c => c.InnerText.Contains(text));
            var count = 0;
            foreach (var openXmlElement in openXmlElements)
            {
                var textRegex = new Regex(text);
                var textMatches = openXmlElement.InnerText.GetMatches(textRegex);
                foreach (var match in textMatches)
                {
                    if (indexes == null || indexes.Any(c => c == count))
                    {
                        var left = openXmlElement.InnerText.Substring(0, match.Index);
                        var leftElement = string.IsNullOrEmpty(left)
                            ? null
                            : openXmlElement.GetClone().SetInnerText(left);
                        var right = openXmlElement.InnerText.Substring(match.Index + 1);
                        var rightElement = string.IsNullOrEmpty(right)
                            ? null
                            : openXmlElement.GetClone().SetInnerText(right);
                        var xmlElement = openXmlElement.GetClone().SetInnermostXml(xml);
                        var replaceList = new List<OpenXmlElement> { leftElement, xmlElement, rightElement };
                        openXmlElement.Replace(replaceList);
                        indexes?.Remove(count);
                        break;
                    }

                    count++;
                }

                if (indexes != null && !indexes.Any())
                {
                    break;
                }
            }

            return element;
        }
        /// <summary>
        /// 将文字替换为文字
        /// </summary>
        /// <param name="element"></param>
        /// <param name="replaces"></param>
        /// <returns></returns>
        public static OpenXmlElement ReplaceInnerText(this OpenXmlElement element, Dictionary<string, string> replaces)
        {
            element.SetInnerText(element.InnerText.ReplaceByDictionary(replaces));
            return element;
        }
    }
}
