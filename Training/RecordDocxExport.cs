using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Docx.Core;
using Training.Entities;

namespace Training
{
    /// <summary>
    /// 自定义Docx的替换规则
    /// </summary>
    public class RecordDocxExport : ExportDocxDefault
    {
        public override string GetValue(object entity, string propName, Dictionary<string, string> attributes)
        {
            if (propName == "Resume.MedicalConsortiumInt")
            {
                //适配Docx中的勾选
                var res = entity.Getter("Resume.MedicalConsortium");
                if (res == null)
                {
                    return "-1";
                }

                return Enum.TryParse<MedicalConsortium>(res.ToString(), true, out var intResult)
                    ? ((int)intResult).ToString()
                    : "-1";
            }

            if (propName == "IsAccommodationInt")
            {
                //适配Docx中的勾选
                var res = entity.Getter("IsAccommodation");
                if (res == null)
                {
                    return "-1";
                }

                return bool.TryParse(res.ToString(), out var boolResult)
                    ? (boolResult ? "1" : "0")
                    : "-1";
            }

            var value = entity.Getter(propName);
            if (value == null)
            {
                return string.Empty;
            }

            string result;
            attributes.TryGetValue("Format", out var format); //格式
            attributes.TryGetValue("Append", out var append); //额外附加的字符串
            if (DateTime.TryParse(value.ToString(), out var time))
            {
                //给时间一个默认Format，省的Docx里面的都加Format
                result = time.ToString(string.IsNullOrEmpty(format) ? "yyyy年MM月dd日" : format,
                    new CultureInfo("zh-chs"));
            }
            else if (propName.EndsWith("Gender") && bool.TryParse(value.ToString(), out var gender))
            {
                result = gender ? "男" : "女";
            }
            else if (propName.EndsWith("Marital"))
            {
                result = ((Marital)value).GetDescription();
            }
            else
            {
                result = value.ToString(format);
            }

            if (string.IsNullOrEmpty(result))
            {
                return result;
            }

            return result + append;
        }
    }
}
