using System;

namespace Training.Entities
{
    /// <summary>
    /// 教育背景
    /// </summary>
    public class Education : EntityBase
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartDate { get; set; }
        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime? EndDate { get; set; }
        /// <summary>
        /// 学校
        /// </summary>
        public string School { get; set; }
        /// <summary>
        /// 学历
        /// </summary>
        public string Name { get; set; }
    }
}
