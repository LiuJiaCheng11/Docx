using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Training
{
    public static class EnumExtension
    {
        /// <summary>
        /// 获取特性DescriptionAttribute的值
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetDescription(this Enum value)
        {
            string str = string.Empty;
            try
            {
                DescriptionAttribute descriptionAttribute =
                    (DescriptionAttribute)((IEnumerable<object>)value.GetType().GetField(value.ToString())
                        .GetCustomAttributes(typeof(DescriptionAttribute), false)).FirstOrDefault<object>();
                str = descriptionAttribute == null ? value.ToString() : descriptionAttribute.Description;
            }
            catch
            {
                // ignore
            }

            return str;
        }
    }
}
