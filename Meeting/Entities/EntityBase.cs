using System;
using System.Collections.Generic;
using System.Text;

namespace Meeting.Entities
{
    /// <summary>
    /// 基类
    /// </summary>
    public abstract class EntityBase
    {
        /// <summary>
        /// 主键
        /// </summary>
        public virtual long Id { get; set; }
    }
}
