using System;
using System.Collections.Generic;
using System.Text;

namespace Meeting.Entities
{
    public class Participant : EntityBase
    {
        /// <summary>
        /// 会议Id
        /// </summary>
        public long ConferenceId { get; set; }
        /// <summary>
        /// 用户Id
        /// </summary>
        public long UserId { get; set; }
        /// <summary>
        /// 用户名
        /// </summary>
        public string UserName { get; set; }
    }
}
