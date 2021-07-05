using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Docx.Entities
{
    public class Student
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public int Duration { get; set; }
        public bool Gender { get; set; }
        public Marital Marital { get; set; }

        public DateTime BirthDay { get; set; }
    }

    public enum MedicalConsortium
    {
        /// <summary>
        /// 否
        /// </summary>
        Null = 0,
        /// <summary>
        /// 医联体单位
        /// </summary>
        MedicalConsortium = 1,
        /// <summary>
        /// 对口支援单位
        /// </summary>
        CounterpartSupport = 2
    }

    /// <summary>
    /// 婚姻状况
    /// </summary>
    public enum Marital
    {
        /// <summary>
        /// 未婚
        /// </summary>
        [Description("未婚")]
        Unmarried = 0,
        /// <summary>
        /// 已婚
        /// </summary>
        [Description("已婚")]
        Married = 1,
        /// <summary>
        /// 丧偶
        /// </summary>
        [Description("丧偶")]
        Widowed = 2,
        /// <summary>
        /// 离婚
        /// </summary>
        [Description("离婚")]
        Divorce = 3
    }


}
