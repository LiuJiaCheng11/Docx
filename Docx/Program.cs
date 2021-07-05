using Docx.Core;
using System;

namespace Docx
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }

    public class Student
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public int Duration { get; set; }
        public bool Gender { get; set; }
        public Marital Marital { get; set; }
    }



    /// <summary>
    /// 自定义Docx的替换规则
    /// </summary>
    public class StudentDocxExport : ExportDocxDefault
    {
        public override string GetValue(object entity, string innerText)
        {
            var value = entity.Getter(innerText);
            return value?.ToString();
        }
    }
}
