using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskUtility
{
    public static class CalendarUtility
    {
        private static DateTime _StartDateofMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        private static DateTime _EndDateofMonth = _StartDateofMonth.AddMonths(1).AddDays(-1);
        /// <summary>
        /// The function will provide the nth working day of the month
        /// </summary>
        /// <returns>number</returns>
        public static int NthBusinessDayOfMonth()
        {
            int Bday = 0;
            DateTime businessDay;
            for (businessDay = DateTime.Today; businessDay >= _StartDateofMonth; businessDay = businessDay.AddDays(-1))
            {
                if (businessDay.DayOfWeek != DayOfWeek.Saturday && businessDay.DayOfWeek != DayOfWeek.Sunday)
                {
                    Bday++;
                }
            }
            return Bday;
        }
        /// <summary>
        /// The function will provide the number of working days remaining in the month
        /// </summary>
        /// <returns>number</returns>
        public static int RemainingBusinessDaysInMonth()
        {
            int Bday = 0;
            DateTime businessDay;
            for (businessDay = DateTime.Today; businessDay <= _EndDateofMonth; businessDay = businessDay.AddDays(1))
            {
                if (businessDay.DayOfWeek != DayOfWeek.Saturday && businessDay.DayOfWeek != DayOfWeek.Sunday)
                {
                    Bday++;
                }
            }
            return Bday;    
        }

    }
}
