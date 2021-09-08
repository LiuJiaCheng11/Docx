using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Docx.Core
{
    public static class DocxExtension
    {
        /// <summary>
        /// 获取Docx文件的字节流
        /// </summary>
        /// <param name="entity">导出需要的信息的实体</param>
        /// <param name="templateFullPath">模板完整路径</param>
        /// <param name="export"></param>
        /// <returns></returns>
        public static byte[] ExportToDocx(this object entity, string templateFullPath, IExportDocx export = null)
        {
            export ??= new ExportDocxDefault(); ////这里应该使用依赖注入获得对应的实现，这里为了简化逻辑，直接就让外面传
            if (!System.IO.File.Exists(templateFullPath))
            {
                return new byte[0];
            }

            var byteArray = File.ReadAllBytes(templateFullPath);
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    var body = wordDoc.MainDocumentPart.Document.Body;
                    FillTable(body, entity, export);
                    FillText(body, entity, export);
                    wordDoc.Save();
                }

                stream.Seek(0, SeekOrigin.Begin);
                var bytes = stream.ToBytes();
                return bytes;
            }
        }

        /// <summary>
        /// 填充文本
        /// </summary>
        /// <param name="body"></param>
        /// <param name="entity"></param>
        /// <param name="export"></param>
        private static void FillText(Body body, object entity, IExportDocx export)
        {
            foreach (var paragraph in body.Descendants<Paragraph>()) //把整个body的内容分成多个段落
            {
                ReplaceParagraph(paragraph, entity, export.GetTextRegex(), export.GetInnerTextRegex(), export);
            }
        }

        /// <summary>
        /// 填充表格
        /// </summary>
        /// <param name="body"></param>
        /// <param name="entity"></param>
        /// <param name="export"></param>
        private static void FillTable(Body body, object entity, IExportDocx export)
        {
            var rowReg = export.GetTableTextRegex();
            var innerRegex = export.GetTableInnerTextRegex();
            foreach (var table in body.Descendants<Table>())
            {
                //在里面替换会退出循环，所以先保存字典，在外面替换
                var replaceRows = new Dictionary<TableRow, List<TableRow>>();
                foreach (var row in table.Descendants<TableRow>())
                {
                    var needRepeat = false; //是否需要重复，行数据才重复
                    //带有{{XX:YY}}的数据的行才会被视为需要处理的行数据，不要带有下标，若是{{XX[0]:YY}}，请使用{XX[0].YY}
                    foreach (var paragraph in row.Descendants<Paragraph>())
                    {
                        var match = rowReg.Match(paragraph.InnerText);
                        if (match.Success)
                        {
                            needRepeat = true;
                            break;
                        }
                    }

                    if (!needRepeat)
                    {
                        continue;
                    }

                    var cells = row.Descendants<TableCell>().ToList();
                    var mergeParas = new List<MergePara>();
                    var column = 1;
                    foreach (var cell in cells)
                    {
                        if (cell.InnerText.HasAttribute("Merge", export.AttributeSplitText, innerRegex))
                        {
                            mergeParas.Add(new MergePara
                            {
                                Column = column
                            });
                        }

                        column++;
                    }

                    var repRows = GetReplaceElements(row, entity, rowReg, innerRegex, export);
                    if (mergeParas.Any())
                    {
                        var isFirst = true;
                        foreach (var repRow in repRows)
                        {
                            column = 1;
                            foreach (var cell in repRow.Descendants<TableCell>())
                            {
                                var mergePara = mergeParas.FirstOrDefault(c => c.Column == column);
                                if (mergePara == null)
                                {
                                    //没有标记要合并，跳过
                                    column++;
                                    continue;
                                }

                                //保证标记了要合并的Element的tcPr子Element必存在
                                var tcPr = cell.PrependChildIfNotExist("<w:tcPr></w:tcPr>".AddNameSpace()
                                    .ToOpenXmlElement());
                                if (mergePara.LastCell != null && mergePara.LastCell.InnerText == cell.InnerText)
                                {
                                    var lastTcPr = mergePara.LastCell.ChildElements.First(c => c.LocalName == "tcPr");
                                    lastTcPr.PrependChildIfNotExist("<w:vAlign w:val = \"center\" />".AddNameSpace()
                                        .ToOpenXmlElement());
                                    lastTcPr.PrependChildIfNotExist("<w:vMerge w:val=\"restart\"/>".AddNameSpace()
                                        .ToOpenXmlElement());
                                    tcPr.PrependChildIfNotExist("<w:vMerge w:val=\"continue\"/>".AddNameSpace()
                                        .ToOpenXmlElement());
                                }
                                else
                                {
                                    //没有合并的，就添加样式居中
                                    tcPr.PrependChildIfNotExist("<w:vAlign w:val = \"center\" />".AddNameSpace()
                                        .ToOpenXmlElement());
                                }

                                mergePara.LastCell = cell;
                                column++;
                            }

                            if (isFirst)
                            {
                                isFirst = false;
                            }
                        }
                    }

                    replaceRows.Add(row, repRows);
                }

                foreach (var replaceRow in replaceRows)
                {
                    replaceRow.Key.Replace(replaceRow.Value);
                }
            }
        }

        private class MergePara
        {
            public int Column { get; set; }

            public TableCell LastCell { get; set; } = null;
        }

        /// <summary>
        /// 替换段落的内容
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="entity"></param>
        /// <param name="regex"></param>
        /// <param name="innerRegex"></param>
        /// <param name="export"></param>
        /// <param name="index">如果数据是数组，则取其第几个元素，仅用于表格</param>
        private static void ReplaceParagraph(OpenXmlElement paragraph, object entity, Regex regex, Regex innerRegex,
            IExportDocx export, int index = -1)
        {
            //拆成最小的段，再实行替换
            var children = paragraph.Descendants<Paragraph>().ToList();
            if (children.Any())
            {
                foreach (var child in children)
                {
                    ReplaceParagraph(child, entity, regex, innerRegex, export, index);
                }

                return;
            }
            //paragraph为1个段落内容，如果是表格的话，则是一个格子的内容

            //获取特性


            if (export.GetTableTextRegex().Match(paragraph.InnerText).Success && index == -1)
            {
                //非表格的重复替换
                var replaceElements = GetReplaceElements(paragraph, entity, export.GetTableTextRegex(),
                    export.GetTableInnerTextRegex(), export);
                paragraph.Replace(replaceElements);
            }
            else if (paragraph.InnerText.Contains("□"))
            {
                //取第一个{}的Text计算值，剩下的作为枚举值，枚举值里都没有则不打勾
                var matches = paragraph.InnerText.GetMatches(innerRegex);
                if (!matches.Any())
                {
                    return;
                }

                var value = entity.GetValue(matches[0].Value, index, export);
                var replaceIndex = -1;
                if (int.TryParse(value, out var intValue))
                {
                    //如果没有{0}{1}{2}至少还会按枚举的值来指示需改变□的序号
                    replaceIndex = intValue;
                }

                var count = 0;
                matches.Remove(matches[0]);
                foreach (var match in matches)
                {
                    if (match.Value == value)
                    {
                        replaceIndex = count;
                    }

                    count++;
                }

                paragraph.SetInnerText(regex.Replace(paragraph.InnerText, ""));
                var wingdings2_0052Xml = "<w:sym w:font=\"Wingdings 2\" w:char=\"0052\" />".AddNameSpace();
                paragraph.ReplaceInnerTextToXml("□", wingdings2_0052Xml, new List<int> { replaceIndex });
            }
            else
            {
                var innerText = paragraph.InnerText.GetReplaceText(entity, regex, innerRegex, export, index);
                paragraph.SetInnerText(innerText);
            }
        }

        /// <summary>
        /// 获取替换后的Text的内容，假定Name的值就是小倉唯，如输入"{Name}こんにちは{Name}！"，则输出"小倉唯こんにちは小倉唯！"
        /// </summary>
        /// <param name="text">文本（可多次出现{Name}）</param>
        /// <param name="entity">数据</param>
        /// <param name="regex">全匹配{Name}的正则</param>
        /// <param name="innerRegex">匹配{Name}里面的Name的正则</param>
        /// <param name="export">导出Docx接口的实现</param>
        /// <param name="index">如果数据是数组，则取其第几个元素，仅用于数组</param>
        private static string GetReplaceText(this string text, object entity, Regex regex, Regex innerRegex, IExportDocx export, int index = -1)
        {
            //找{XXX},{YYY}...matches的顺序是按出现的先后顺序的
            //★不能使用Replace，因为替换的{XXX}的值，可能出现{YYY}
            var matches = text.GetMatches(regex);
            //这个段落没有需要替换的，跳过
            if (!matches.Any()) return text;

            //allText的左部
            var leftAllText = string.Empty;
            //allText的右部
            var rightAllText = text;
            //要减去的index
            var minusIndex = 0;
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                var matchText = match.Value; //形如 {XXX}
                var innerText = matchText.FirstMatchOrDefault(innerRegex);
                var value = entity.GetValue(innerText, index, export);
                if (value == innerText)
                {
                    //不予替换
                    continue;
                }

                //对右部切分
                var actualIndex = match.Index - minusIndex; //匹配的字符串对于右边来说的Index
                minusIndex = match.Index + match.Length;
                //把match.Index之前的 + value切给左部
                leftAllText += rightAllText.Substring(0, actualIndex) + value;
                //把match.Index + length往后的字符串切给右部
                rightAllText = rightAllText.Substring(actualIndex + match.Length,
                    rightAllText.Length - (actualIndex + match.Length));
            }

            text = leftAllText + rightAllText;
            return text;
        }

        /// <summary>
        /// 获取重复替换的OpenXmlElement
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="openXmlElement"></param>
        /// <param name="entity"></param>
        /// <param name="regex"></param>
        /// <param name="innerRegex"></param>
        /// <param name="export"></param>
        /// <returns></returns>
        private static List<T> GetReplaceElements<T>(T openXmlElement, object entity, Regex regex, Regex innerRegex,
            IExportDocx export) where T : OpenXmlElement
        {
            //数据行，循环次数为所有行数据的最小值，如混用时，XX数组长度为2，YY数组长度为3，那么循环到第3行，报错退出
            var index = 0;
            var replaces = new List<T>();
            while (index < 999) //防止死循环，如{{XX[0]:YY}}没有用上index，就不会抛异常
            {
                var clone = (T)openXmlElement.Clone();
                try
                {
                    ReplaceParagraph(clone, entity, regex, innerRegex, export, index);
                }
                catch (Exception ex)
                {
                    if (ex is ArgumentOutOfRangeException)
                    {
                        //超出数组长度，报错退出
                        break;
                    }

                    //其他异常
                    throw;
                }

                replaces.Add(clone);
                index++;
            }

            return replaces;
        }

        /// <summary>
        /// 将流转成字节数组
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        private static byte[] ToBytes(this Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }

        /// <summary>
        /// 获取数组中的元素
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="arrayProp"></param>
        /// <param name="index"></param>
        /// <param name="export"></param>
        /// <param name="filterDic"></param>
        /// <returns></returns>
        public static object GetArrayItem(this object entity, string arrayProp, int index, IExportDocx export,
            Dictionary<string, string> filterDic = null)
        {
            var rangeEntities = (IEnumerable<object>)entity.Getter(arrayProp);
            if (rangeEntities == null || index <= -1)
            {
                throw new ArgumentOutOfRangeException();
            }

            if (filterDic != null && filterDic.Any())
            {
                rangeEntities = rangeEntities.Where(rangeEntity =>
                {
                    foreach (var kvp in filterDic)
                    {
                        var val = rangeEntity.GetValue(kvp.Key, -1, export);
                        if (val != kvp.Value)
                        {
                            return false;
                        }
                    }

                    return true;
                });
            }

            //这里超出数组会抛异常
            var result = rangeEntities.ElementAt(index);
            return result;
        }

        /// <summary>
        /// 获取值（兼容InnerText是XX:YY|ZZ|AA=??形式）
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="innerText"></param>
        /// <param name="index"></param>
        /// <param name="export"></param>
        /// <returns></returns>
        private static string GetValue(this object entity, string innerText, int index, IExportDocx export)
        {
            var value = innerText;
            var field = innerText.GetFirstStr(export.AttributeSplitText);
            if (field.Contains(export.RangeSplitText))
            {
                var array = field.Split(export.RangeSplitText);
                if (array.Length != 2)
                {
                    //错误语法，不予替换
                    return value;
                }

                var attributes = innerText.GetAttributes(export.AttributeSplitText);
                var filterDic = attributes.GetFilterDic();
                var arrayItemEntity = entity.GetArrayItem(array[0], index, export, filterDic);
                value = export.GetValue(arrayItemEntity, array[1]);
            }
            else
            {
                value = export.GetValue(entity, field);
            }

            return value;
        }

        /// <summary>
        /// 获取字符串分割后的第一个字符串
        /// </summary>
        /// <param name="source"></param>
        /// <param name="splitStr"></param>
        /// <returns></returns>
        private static string GetFirstStr(this string source, string splitStr)
        {
            return string.IsNullOrEmpty(source) ? string.Empty : source.Split(splitStr)[0];
        }
        /// <summary>
        /// 检测是否含有某个属性（只要有一个含有该属性就行{A|属性}{B}）
        /// </summary>
        /// <param name="text"></param>
        /// <param name="attributeText"></param>
        /// <param name="attributeSplitText"></param>
        /// <param name="innerTextRegex"></param>
        /// <returns></returns>
        private static bool HasAttribute(this string text, string attributeText, string attributeSplitText, Regex innerTextRegex)
        {
            var innerTexts = text.GetMatches(innerTextRegex);
            foreach (var innerText in innerTexts)
            {
                var attributes = innerText.Value.GetAttributes(attributeSplitText);
                if (attributes.Any(c => c == attributeText))
                {
                    return true;
                }
            }

            return false;
        }
        /// <summary>
        /// 获取所有属性
        /// </summary>
        /// <param name="innerText"></param>
        /// <param name="attributeSplitText"></param>
        /// <returns></returns>
        private static List<string> GetAttributes(this string innerText, string attributeSplitText)
        {
            var array = innerText.Split(attributeSplitText);
            var result = new List<string>();
            for (int i = 1; i < array.Length; i++)
            {
                result.Add(array[i]);
            }

            return result;
        }
        /// <summary>
        /// 获取筛选数组的字典
        /// </summary>
        /// <param name="attributes"></param>
        /// <returns></returns>
        private static Dictionary<string, string> GetFilterDic(this List<string> attributes)
        {
            var filterDic = new Dictionary<string, string>();
            foreach (var attribute in attributes)
            {
                if (!attribute.Contains("="))
                {
                    continue;
                }

                var arr = attribute.Split("=");
                if (arr.Length == 0)
                {
                    continue;
                }

                var key = arr[0].Trim();
                if (filterDic.ContainsKey(key))
                {
                    continue;
                }

                var val = string.Empty;
                if (arr.Length > 1)
                {
                    val = arr[1].Trim();
                }

                filterDic.Add(key, val);
            }

            return filterDic;
        }
    }
}
