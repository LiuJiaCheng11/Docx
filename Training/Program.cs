using Docx.Core;
using System;
using System.Collections.Generic;
using System.IO;
using Training.Entities;

namespace Training //教学进修
{
    class Program
    {
        static void Main(string[] args)
        {
            //因为复制到输出目录，所以直接在输出目录下找
            var templateFullPath = System.Environment.CurrentDirectory + "/Templates/RecordExport.docx"; 
            var n = DateTime.Now;
            var outputDirectory =
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}";
            var directoryInfo = new DirectoryInfo(outputDirectory);
            directoryInfo.Create();
            var outputPath = $"{directoryInfo.FullName}/RecordExport.docx";
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
                    StartingDateOfFirstJob = new DateTime(2020, 1, 1),
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

            //应使用ModCode注册IDocxExport的实现为RecordDocxExport，由于代码里为了简化逻辑，直接就new了一个，所以这里省略注册
            var bytes = record.ExportToDocx(templateFullPath, new RecordDocxExport());
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
}
