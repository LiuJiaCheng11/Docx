using System;
using System.Collections.Generic;
using System.Text;

namespace Meeting.Entities
{
    /// <summary>
    /// 议题
    /// </summary>
    public class Agenda : EntityBase
    {
        /// <summary>
        /// 议题名称
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 主讲人Id
        /// </summary>
        public long SpeakerId { get; set; }
        /// <summary>
        /// 主讲人
        /// </summary>
        public string Speaker { get; set; }

        /// <summary>
        /// 时长（分）
        /// </summary>
        public int Duration { get; set; } = 10;
        /// <summary>
        /// 序号
        /// </summary>
        public int Index { get; set; }
        /// <summary>
        /// 分组
        /// </summary>
        public string Group { get; set; }
    }
}
