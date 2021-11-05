using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Docx.Core
{
    /// <summary>
    /// 默认格式参数
    /// </summary>
    public class DefaultFormatter : IFormatProvider, ICustomFormatter
    {
        public object GetFormat(Type formatType)
        {
            return formatType == typeof(ICustomFormatter) ? this : null;
        }

        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            try
            {
                return HandleOtherFormats(format, arg);
            }
            catch (FormatException)
            {
                return arg.ToString();
            }
        }

        private string HandleOtherFormats(string format, object arg)
        {
            string result;
            if (arg is IFormattable)
            {
                format ??= string.Empty;
                if ((format == "中文" || format.ToLower() == "chinese") && long.TryParse(arg.ToString(), out var numeric))
                {
                    result = numeric.ToChinese();
                }
                else
                {
                    result = ((IFormattable)arg).ToString(format, CultureInfo.CurrentCulture);
                }
            }
            else if (arg != null)
                result = arg.ToString();
            else
                result = string.Empty;
            return result;
        }
    }
}
