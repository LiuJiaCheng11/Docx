using Docx.Core;
using Docx.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Docx
{
    /* 导出Docx的调用方 */
    class Program
    {
        static void Main(string[] args)
        {
            //应使用ModCode注册IDocxExport的实现为RecordDocxExport，由于代码里为了简化逻辑，直接就new了一个，所以这里省略注册
            var templateFullPath = System.Environment.CurrentDirectory + "/Templates/RecordExport.docx"; //因为复制到输出目录，所以直接在目录下找
            var outputPath = System.Environment.CurrentDirectory + "/RecordExport.docx";
            var record = new Record
            {
                AbilityRemark = "a",
                Comment = "c",
                Duration = 1,
                Id = 1,
                IsAccommodation = true,
                Remark = "r",
                Request = "r",
                TrainingSubject = "s",
                Resume = new Resume
                {
                    BirthDay = new DateTime(2000, 1, 1),
                    Gender = false,
                    Hospital = "h",
                    HospitalDepartment = "hd",
                    Id = 2,
                    Marital = Marital.Divorce,
                    MedicalConsortium = MedicalConsortium.CounterpartSupport,
                    Name = "n",
                    PoliticalStatus = "ps",
                    PracticeCertificate = "pc",
                    QualificationCertificate = "qc",
                    StartingDateOfFirstJob = new DateTime(2020,1,1),
                    Title = "t",
                    TitleConferredDate = new DateTime(2021, 1, 1),
                    Educations = new List<Education>
                    {
                        new Education
                        {
                            StartDate = new DateTime(2015, 9, 1),
                            EndDate = new DateTime(2018, 7, 1),
                            School = "XX高中",
                            Name = "高中"
                        },
                        new Education
                        {
                            StartDate = new DateTime(2018, 9, 1),
                            EndDate = new DateTime(2022, 7, 1),
                            School = "XX大学",
                            Name = "本科"
                        },
                        new Education
                        {
                            StartDate = new DateTime(2022, 9, 1),
                            EndDate = new DateTime(2025, 7, 1),
                            School = "YY大学",
                            Name = "硕士"
                        }
                    },
                    Experiences = new List<Experience>
                    {
                        new Experience
                        {
                            StartDate = new DateTime(2025, 8, 1),
                            EndDate = null,
                            Company = "XX医院",
                            Position = "规培生"
                        }
                    }
                }
            };

            var bytes = record.ExportToDocx(templateFullPath, "123");
            if (System.IO.File.Exists(outputPath))
            {
                System.IO.File.Delete(outputPath);
            }

            using (var fs = new System.IO.FileStream(outputPath, System.IO.FileMode.Create))
            {
                fs.Write(bytes, 0, bytes.Length);
            }

            Console.WriteLine("Hello World!");
        }
    }

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

    public static class EnumExtension
    {
        /// <summary>
        /// 获取特性DescriptionAttribute的值
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetDescription(this Enum value)
        {
            string str = string.Empty;
            try
            {
                DescriptionAttribute descriptionAttribute = (DescriptionAttribute)((IEnumerable<object>)value.GetType().GetField(value.ToString()).GetCustomAttributes(typeof(DescriptionAttribute), false)).FirstOrDefault<object>();
                str = descriptionAttribute == null ? value.ToString() : descriptionAttribute.Description;
            }
            catch
            {
            }
            return str;
        }
    }
}
