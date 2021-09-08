using System;
using System.Collections.Generic;
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
        public override string GetValue(object entity, string propName)
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

            var result = value.ToString();
            if (propName.EndsWith("Gender") && bool.TryParse(result, out var gender))
            {
                result = gender ? "男" : "女";
            }
            else if (propName.EndsWith("Marital") && int.TryParse(result, out var marital))
            {
                result = ((Marital)marital).GetDescription();
            }
            else if (DateTime.TryParse(result, out var time))
            {
                result = time.ToString("yyyy年MM月dd日");
            }

            return result;
        }
    }
}
