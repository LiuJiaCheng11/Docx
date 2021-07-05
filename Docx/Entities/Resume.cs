using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Docx.Entities
{
    /// <summary>
    /// 进修档案
    /// </summary>
    public class Resume : EntityBase
    {
        /// <summary>
        /// 医联体类型
        /// </summary>
        public MedicalConsortium MedicalConsortium { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 性别
        /// </summary>
        public bool Gender { get; set; }
        /// <summary>
        /// 婚姻状况
        /// </summary>
        public Marital Marital { get; set; }
        /// <summary>
        /// 出生年月
        /// </summary>
        public DateTime BirthDay { get; set; }
        /// <summary>
        /// 参加工作时间
        /// </summary>
        public DateTime? StartingDateOfFirstJob { get; set; }
        /// <summary>
        /// 政治面目
        /// </summary>
        public string PoliticalStatus { get; set; }
        /// <summary>
        /// 技术职称
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 评定时间
        /// </summary>
        public DateTime? TitleConferredDate { get; set; }
        /// <summary>
        /// 工作单位
        /// </summary>
        public string Hospital { get; set; }
        /// <summary>
        /// 工作科室
        /// </summary>
        public string HospitalDepartment { get; set; }
        /// <summary>
        /// 医师资格证书编号
        /// </summary>
        public string QualificationCertificate { get; set; }
        /// <summary>
        /// 医师（护士）执业证书编号
        /// </summary>
        public string PracticeCertificate { get; set; }
        /// <summary>
        /// 教育背景
        /// </summary>
        public List<Education> Educations { get; set; }
        /// <summary>
        /// 工作经历
        /// </summary>
        public List<Experience> Experiences { get; set; }
    }

    /// <summary>
    /// 医联体类型
    /// </summary>
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
