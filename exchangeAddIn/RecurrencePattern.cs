using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Radvision.Scopia.ExchangeMeetingAddIn
{
    public class RecurrencePattern
    {
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public bool hasEnd { get; set; }
        public int numberOfOccurrences { get; set; }
        public PatternType patternType { get; set; }
        public int dailyInterval { get; set; }
        public int weeklyInterval { get; set; }
        public int monthlyInterval { get; set; }
        //By the fixed order from Sunday to Monday 
        public icm.XmlApi.dayOfWeekType[] daysOfWeek { get; set; }
        public int dayOfMonth { get; set; }
        public icm.XmlApi.WeekOfMonthType weekOfMonth { get; set; }
        public icm.XmlApi.dayRecurrenceType dayOfWeekOfMonthly { get; set; }

        public RecurrencePattern() {
            daysOfWeek = null;
        }

        public string getStringForHash() {
            StringBuilder hashString = new StringBuilder("");
            hashString.Append(startDate.ToUniversalTime())
            .Append(endDate.ToUniversalTime())
            .Append(hasEnd)
            .Append(numberOfOccurrences)
            .Append(patternType)
            .Append(dailyInterval)
            .Append(weeklyInterval)
            .Append(monthlyInterval)
            .Append(dayOfMonth)
            .Append(weekOfMonth)
            .Append(dayOfWeekOfMonthly);

            if(null != daysOfWeek)
                foreach (icm.XmlApi.dayOfWeekType dayOfWeek in daysOfWeek) {
                    hashString.Append(dayOfWeek);
                }

            return hashString.ToString();
        }
    }

    public enum PatternType { 
        DAILY,
        WEEKLY,
        MONTHLY,
        RELATIVE_MONTHLY,
        YEARLY,
        RELATIVE_YEARLY
    }
}
