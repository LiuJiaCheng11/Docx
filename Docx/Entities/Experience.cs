using System;
using System.Collections.Generic;
using System.Text;

namespace Docx.Entities
{
    /// <summary>
    /// 工作经历
    /// </summary>
    public class Experience : EntityBase
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
        /// 工作单位及部门
        /// </summary>
        public string Company { get; set; }
        /// <summary>
        /// 职务
        /// </summary>
        public string Position { get; set; }
    }
}
