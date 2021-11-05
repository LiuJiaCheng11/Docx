using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Docx.Core;
using Meeting.Entities;

namespace Meeting
{
    public class ConferenceDocxExport : ExportDocxDefault
    {
        public override string GetValue(object entity, string propName, Dictionary<string, string> attributes)
        {
            var value = entity.Getter(propName);
            if (value == null)
            {
                return string.Empty;
            }

            var result = string.Empty;
            attributes.TryGetValue("Format", out var format); //格式
            attributes.TryGetValue("Append", out var append); //额外附加的字符串
            if (DateTime.TryParse(value.ToString(), out var time))
            {
                result = time.ToString(string.IsNullOrEmpty(format) ? "yyyy年MM月dd日（dddd） HH:mm" : format,
                    new CultureInfo("zh-chs"));
                if (append == "上下午")
                {
                    append = time.Hour >= 12 ? " 下午" : " 上午";
                }
            }
            else if (propName == "Participants")
            {
                try
                {
                    var participants = (IEnumerable<Participant>)value;
                    result = string.Join("、", participants.Select(c => c.UserName));
                }
                catch
                {
                    // ignore
                }
            }
            else if (propName == "Index")
            {
                if (int.TryParse(value.ToString(), out var index) && index > 0)
                {
                    //排除0
                    result = value.ToString(format);
                }
            }
            else
            {
                return value.ToString(format);
            }

            if (string.IsNullOrEmpty(result))
            {
                return result;
            }

            return result + append;
        }
    }
}
