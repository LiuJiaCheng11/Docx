using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Docx.Core
{
    /// <summary>
    /// 导出Docx文件的参数
    /// </summary>
    /// <summary>
    /// 导出Docx文件的参数
    /// </summary>
    public interface IExportDocx
    {
        /// <summary>
        /// 表示此处应重复替换，并换行显示，此符号前面的为数组，此符号后面的为数组的字段，形如{{XX:YY}}
        /// </summary>
        string RangeSplitText { get; set; }
        /// <summary>
        /// 表示此处有特性，
        /// </summary>
        string AttributeSplitText { get; set; }
        /// <summary>
        /// 匹配普通文本
        /// </summary>
        string TextLeft { get; set; }
        /// <summary>
        /// 匹配普通文本
        /// </summary>
        string TextRight { get; set; }
        /// <summary>
        /// 匹配表格文本
        /// </summary>
        string TableTextLeft { get; set; }
        /// <summary>
        /// 匹配表格文本
        /// </summary>
        string TableTextRight { get; set; }
        /// <summary>
        /// 获取匹配的正则表达式，一般不需要重写
        /// </summary>
        /// <returns></returns>
        Regex GetTextRegex();
        /// <summary>
        /// 获取匹配的内容的正则表达式，一般不需要重写
        /// </summary>
        /// <returns></returns>
        Regex GetInnerTextRegex();
        /// <summary>
        /// 获取匹配的正则表达式，一般不需要重写
        /// </summary>
        /// <returns></returns>
        Regex GetTableTextRegex();
        /// <summary>
        /// 获取匹配的内容的正则表达式，一般不需要重写
        /// </summary>
        /// <returns></returns>
        Regex GetTableInnerTextRegex();

        /// <summary>
        /// 字段输出值
        /// </summary>
        /// <param name="entity">数据模型</param>
        /// <param name="propName">字段名</param>
        /// <param name="attributes">导出的属性</param>
        /// <returns>替换后的内容</returns>
        string GetValue(object entity, string propName, Dictionary<string, string> attributes);
    }

    /// <summary>
    /// 导出Docx文件的参数（默认实现）
    /// </summary>
    public class ExportDocxDefault : IExportDocx
    {
        public virtual string RangeSplitText { get; set; } = ":";
        public string AttributeSplitText { get; set; } = "|";
        public virtual string TextLeft { get; set; } = "{";
        public virtual string TextRight { get; set; } = "}";
        public virtual string TableTextLeft { get; set; } = "{{";
        public virtual string TableTextRight { get; set; } = "}}";

        protected Regex TextRegex;
        public virtual Regex GetTextRegex()
        {
            if (TextRegex != null)
            {
                return TextRegex;
            }

            var left = TextLeft.GetPattern();
            var right = TextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = "[^" + left + right + rangeSplitText + "]*?";
            var pattern = left + inner + right;
            TextRegex = new Regex(pattern);
            return TextRegex;
        }

        protected Regex InnerTextRegex;
        public virtual Regex GetInnerTextRegex()
        {
            if (InnerTextRegex != null)
            {
                return InnerTextRegex;
            }

            var left = TextLeft.GetPattern();
            var right = TextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = "[^" + left + right + rangeSplitText + "]*?";
            var pattern = "(?<=" + left + ")(" + inner + ")(?=" + right + ")";
            InnerTextRegex = new Regex(pattern);
            return InnerTextRegex;
        }

        protected Regex TableTextRegex;
        public virtual Regex GetTableTextRegex()
        {
            if (TableTextRegex != null)
            {
                return TableTextRegex;
            }

            var left = TableTextLeft.GetPattern();
            var right = TableTextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = ".*?" + rangeSplitText + ".*?";
            var pattern = left + inner + right;
            TableTextRegex = new Regex(pattern);
            return TableTextRegex;
        }

        protected Regex TableInnerTextRegex;
        public virtual Regex GetTableInnerTextRegex()
        {
            if (TableInnerTextRegex != null)
            {
                return TableInnerTextRegex;
            }

            var left = TableTextLeft.GetPattern();
            var right = TableTextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = ".*?" + rangeSplitText + ".*?";
            var pattern = "(?<=" + left + ")(" + inner + ")(?=" + right + ")";
            TableInnerTextRegex = new Regex(pattern);
            return TableInnerTextRegex;
        }

        public virtual string GetValue(object entity, string propName, Dictionary<string, string> attributes)
        {
            var value = entity.Getter(propName);
            attributes.TryGetValue("Format", out var format); //格式
            attributes.TryGetValue("Append", out var append); //额外附加的字符串
            var result = value?.ToString(format);
            if (string.IsNullOrEmpty(result))
            {
                return result;
            }

            return result + append;
        }
    }
}
