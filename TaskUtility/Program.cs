using System;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace TaskUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            Parallel.ForEach(XElement.Load(@"AppData\BusinessConfiguration.xml").Elements("Business"), (XElement level1Element) => 
            {
                var referenceDate = Int32.Parse(level1Element.Attribute("remainingWorkDays").Value);//will be populated from xml varies for region
                var countryId = Int32.Parse(level1Element.Attribute("id").Value);
                var remBusinessDays = CalendarUtility.RemainingBusinessDaysInMonth();
                var nthBusinessDay = CalendarUtility.RemainingBusinessDaysInMonth() == referenceDate ? CalendarUtility.RemainingBusinessDaysInMonth() : CalendarUtility.NthBusinessDayOfMonth();
                var isFromEnd = remBusinessDays == referenceDate ? true : false;
                var bc = new BusinessConfiguration(countryId, nthBusinessDay, isFromEnd);
                //Execute the pending tasks from previous days business rules
                CarryForwardBusinessConfiguration carryConfig = new CarryForwardBusinessConfiguration(bc.CountryId);
                if (bc.RuleId!=0? bc.ExecuteBusinessRule(bc):false)
                {
                    carryConfig = new CarryForwardBusinessConfiguration(bc);
                }
            });
            Console.ReadKey();
        }

        
    }
}
