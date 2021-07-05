using System;
using System.Collections.Generic;
using System.Text;

namespace Docx.Entities
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
