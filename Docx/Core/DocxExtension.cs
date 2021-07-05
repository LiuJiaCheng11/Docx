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
        /// <param name="modCode">导出配置接口注入时的ModCode</param>
        /// <returns></returns>
        public static byte[] ExportToDocx(this object entity, string templateFullPath, string modCode = "")
        {
            var export = string.IsNullOrEmpty(modCode)
                ? new ExportDocxDefault()
                : new RecordDocxExport(); //这里应该使用依赖注入获得对应的实现，这里为了简化逻辑，直接就new了一个
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
                    FillText(body, entity, export);
                    FillTable(body, entity, export);
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
                ReplaceParagraphText(paragraph, entity, export.GetTextRegex(), export.GetInnerTextRegex(), export);
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

                    //数据行，循环次数为所有行数据的最小值，如混用时，XX数组长度为2，YY数组长度为3，那么循环到第3行，报错退出
                    var index = 0;
                    var delRow = row;
                    var addRows = new List<TableRow>();
                    while (index < 999) //防止死循环，如{{XX[0]:YY}}没有用上index，就不会抛异常
                    {
                        var cloneRow = (TableRow)row.Clone();
                        try
                        {
                            foreach (var paragraph in cloneRow.Descendants<Paragraph>())
                            {
                                ReplaceParagraphText(paragraph, entity, rowReg, innerRegex, export, index);
                            }
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

                        addRows.Add(cloneRow);
                        index++;
                    }

                    replaceRows.Add(delRow, addRows);
                }

                foreach (var replaceRow in replaceRows)
                {
                    replaceRow.Key.Replace(replaceRow.Value);
                }
            }
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
        private static void ReplaceParagraphText(OpenXmlElement paragraph, object entity, Regex regex, Regex innerRegex,
            IExportDocx export, int index = -1)
        {
            //paragraph为1个段落内容，如果是表格的话，则是一个格子的内容
            
            //不确定paragraph.InnerText(string)的内容是否就是断句后的Text(string)的相加后的结果，所以先不用它，用自己组装的allText
            if (paragraph.InnerText.Contains("□"))
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
                var wingdings2_0052Xml =
                    "<w:sym xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:font=\"Wingdings 2\" w:char=\"0052\" />";
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
        /// <param name="index">如果数据是数组，则取其第几个元素，仅用于表格</param>
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
        /// <returns></returns>
        public static object GetArrayItem(this object entity, string arrayProp, int index)
        {
            var rangeEntities = (IEnumerable<object>)entity.Getter(arrayProp);
            if (rangeEntities == null || index <= -1)
            {
                throw new ArgumentOutOfRangeException();
            }

            //这里超出数组会抛异常
            var result = rangeEntities.ElementAt(index);
            return result;
        }

        /// <summary>
        /// 获取值（兼容InnerText是XX:YY形式）
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="innerText"></param>
        /// <param name="index"></param>
        /// <param name="export"></param>
        /// <returns></returns>
        private static string GetValue(this object entity, string innerText, int index, IExportDocx export)
        {
            var value = innerText;
            if (innerText.Contains(export.RangeSplitText))
            {
                var array = innerText.Split(export.RangeSplitText);
                if (array.Length != 2)
                {
                    //错误语法，不予替换
                    return value;
                }

                var arrayItemEntity = entity.GetArrayItem(array[0], index);
                value = export.GetValue(arrayItemEntity, array[1]);
            }
            else
            {
                value = export.GetValue(entity, innerText);
            }

            return value;
        }
    }
}
