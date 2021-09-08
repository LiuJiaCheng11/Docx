namespace Training.Entities
{
    /// <summary>
    /// 进修记录
    /// </summary>
    public class Record : EntityBase
    {
        /// <summary>
        /// 进修科目
        /// </summary>
        public string TrainingSubject { get; set; }
        /// <summary>
        /// 进修期限（单位：月）
        /// </summary>
        public int Duration { get; set; }
        /// <summary>
        /// 是否住宿
        /// </summary>
        public bool IsAccommodation { get; set; }
        /// <summary>
        /// 进修目的和要求
        /// </summary>
        public string Request { get; set; }
        /// <summary>
        /// 业务能力掌握情况
        /// </summary>
        public string AbilityRemark { get; set; }
        /// <summary>
        /// 进修意见
        /// </summary>
        public string Comment { get; set; }
        /// <summary>
        /// 备注
        /// </summary>
        public string Remark { get; set; }
        /// <summary>
        /// 进修档案
        /// </summary>
        public Resume Resume { get; set; }
    }
}
