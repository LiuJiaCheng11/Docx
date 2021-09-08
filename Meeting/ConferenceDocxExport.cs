using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Docx.Core;
using Meeting.Entities;

namespace Meeting
{
    public class ConferenceDocxExport : ExportDocxDefault
    {
        public override string GetValue(object entity, string propName)
        {
            var value = entity.Getter(propName);
            if (value == null)
            {
                return string.Empty;
            }

            var result = string.Empty;
            if (DateTime.TryParse(value.ToString(), out var time))
            {
                result = time.ToString("yyyy年MM月dd日（dddd） HH:mm");
            }
            else if (propName == "Participants")
            {
                try
                {
                    var participants = (IEnumerable<Participant>)value;
                    result = string.Join("、", participants.Select(c => c.UserName));
                }
                catch
                {
                    // ignore
                }
            }
            else
            {
                result = value.ToString();
            }

            return result;
        }
    }
}
