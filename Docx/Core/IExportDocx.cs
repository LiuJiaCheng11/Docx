using System.Text.RegularExpressions;

namespace Docx.Core
{
    /// <summary>
    /// 导出Docx文件的参数
    /// </summary>
    public interface IExportDocx
    {
        /// <summary>
        /// 表示此处应重复替换，并换行显示，此符号前面的为数组，此符号后面的为数组的字段，形如{}
        /// </summary>
        string RangeSplitText { get; set; }
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
        /// 替换特殊标签
        /// </summary>
        /// <param name="entity">数据模型</param>
        /// <param name="propName">innerText</param>
        /// <returns>替换后的内容</returns>
        string GetValue(object entity, string propName);
    }

    /// <summary>
    /// 导出Docx文件的参数（默认实现）
    /// </summary>
    public class ExportDocxDefault : IExportDocx
    {
        public virtual string RangeSplitText { get; set; } = ":";

        public virtual string TextLeft { get; set; } = "{";
        public virtual string TextRight { get; set; } = "}";
        public virtual string TableTextLeft { get; set; } = "{{";
        public virtual string TableTextRight { get; set; } = "}}";

        public virtual Regex GetTextRegex()
        {
            var left = TextLeft.GetPattern();
            var right = TextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = "[^" + left + right + rangeSplitText + "]*?";
            var pattern = left + inner + right;
            return new Regex(pattern);
        }

        public virtual Regex GetInnerTextRegex()
        {
            var left = TextLeft.GetPattern();
            var right = TextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = "[^" + left + right + rangeSplitText + "]*?";
            var pattern = "(?<=" + left + ")(" + inner + ")(?=" + right + ")";
            return new Regex(pattern);
        }

        public virtual Regex GetTableTextRegex()
        {
            var left = TableTextLeft.GetPattern();
            var right = TableTextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = ".*?" + rangeSplitText + ".*?";
            var pattern = left + inner + right;
            return new Regex(pattern);
        }

        public virtual Regex GetTableInnerTextRegex()
        {
            var left = TableTextLeft.GetPattern();
            var right = TableTextRight.GetPattern();
            var rangeSplitText = RangeSplitText.GetPattern();
            var inner = ".*?" + rangeSplitText + ".*?";
            var pattern = "(?<=" + left + ")(" + inner + ")(?=" + right + ")";
            return new Regex(pattern);
        }

        public virtual string GetValue(object entity, string propName)
        {
            var value = entity.Getter(propName);
            return value?.ToString();
        }
    }
}
