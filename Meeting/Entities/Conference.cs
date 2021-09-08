using System;
using System.Collections.Generic;
using System.Text;

namespace Meeting.Entities
{
    public class Conference : EntityBase
    {
        /// <summary>
        /// 会议室
        /// </summary>
        public string Room { get; set; }

        /// <summary>
        /// 主题
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 主持人
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// 参会人员
        /// </summary>
        public List<Participant> Participants { get; set; }

        /// <summary>
        /// 会议议题
        /// </summary>
        public List<Agenda> Agendas { get; set; }
    }
}
