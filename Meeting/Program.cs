using System;
using System.Collections.Generic;
using System.IO;
using Docx.Core;
using Meeting.Entities;

namespace Meeting
{
    class Program
    {
        static void Main(string[] args)
        {
            //因为复制到输出目录，所以直接在输出目录下找
            var templateFullPath = System.Environment.CurrentDirectory + "/Templates/ConferenceExport.docx";
            var n = DateTime.Now;
            var outputDirectory =
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}";
            var directoryInfo = new DirectoryInfo(outputDirectory);
            directoryInfo.Create();
            var outputPath = $"{directoryInfo.FullName}/ConferenceExport.docx";
            var conference = new Conference
            {
                Subject = "庆祝中国共产党成立100周年大会",
                Room = "4楼会议室",
                StartTime = new DateTime(2021, 7, 1, 8, 0, 0),
                EndTime = new DateTime(2021, 7, 1, 18, 0, 0),
                Host = "张三",
                Participants = new List<Participant>
                {
                    new Participant
                    {
                        UserName = "张三"
                    },
                    new Participant
                    {
                        UserName = "李四"
                    },
                    new Participant
                    {
                        UserName = "王五"
                    },
                    new Participant
                    {
                        UserName = "赵六"
                    }
                },
                Agendas = new List<Agenda>
                {
                    new Agenda
                    {
                        Index = 1,
                        Title = "议题A",
                        Group = "学习",
                        Speaker = "张三"
                    },
                    new Agenda
                    {
                        Index = 2,
                        Title = "议题B",
                        Group = "学习",
                        Speaker = "李四"
                    },
                    new Agenda
                    {
                        Index = 3,
                        Title = "议题C",
                        Group = "学习",
                        Speaker = "李四",
                        Duration = 15
                    },
                    new Agenda
                    {
                        Index = 4,
                        Title = "议题D",
                        Group = "学习",
                        Speaker = "王五",
                        Duration = 10
                    },
                    new Agenda
                    {
                        Index = 5,
                        Title = "议题E",
                        Group = "审定",
                        Speaker = "王五",
                        Duration = 10
                    },
                    new Agenda
                    {
                        Index = 6,
                        Title = "议题F",
                        Group = "审定",
                        Speaker = "王五",
                        Duration = 10
                    },
                    new Agenda
                    {
                        Index = 7,
                        Title = "议题G",
                        Group = "通传",
                        Speaker = "张三",
                        Duration = 10
                    },
                }
            };

            //应使用ModCode注册IDocxExport的实现为RecordDocxExport，由于代码里为了简化逻辑，直接就new了一个，所以这里省略注册
            var bytes = conference.ExportToDocx(templateFullPath, new ConferenceDocxExport());
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
