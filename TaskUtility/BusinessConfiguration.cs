using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using TaskUtility.DataAccess;

namespace TaskUtility
{
    public delegate void ExecuteTask(BusinessConfiguration rule);
    public class BusinessConfiguration : IBusinessConfiguration
    {
        public BusinessConfiguration() { }
        public BusinessConfiguration(int countryId, int dayOfTheMonth, bool isFromEnd) {
            //Execute the rules when the object is created
            GetBusinessRule(countryId, dayOfTheMonth, isFromEnd);
        }
        public long RuleId { get; set; }

        public int CountryId { get; set; }
        public int NthWorkingDayOfMonth { get; set; }
        
        public int RemainingWorkingDaysOfMonth { get; set; }
        
        [DataType(DataType.EmailAddress)]
        public string Recipient { get; set; }
        [DataType(DataType.EmailAddress)]
        public string Owner { get; set; }

        public string Rule { get; set; }

        public Freequency RuleFreequency { get; set; }
        
        private BusinessConfiguration GetBusinessRule(int countryId, int dayOfTheMonth, bool isFromEnd)
        {
            CountryId = countryId;
            //connect to the database and fetch the list of Business rules greater than the nth day
            using (TaskUtilityAutomationContainer context = new TaskUtilityAutomationContainer())
            {
                var bc = (from businessConfig in context.Rules
                      join userCountryRole in context.UserCountryRoles on businessConfig.UserCountryRole_Id equals userCountryRole.Id
                      join owner in context.Users on userCountryRole.User_Id equals owner.Id
                      join userCountryRole1 in context.UserCountryRoles on businessConfig.UserCountryRole_Id equals userCountryRole1.Id
                      join recipient in context.Users on userCountryRole1.User_Id equals recipient.Id
                      join country in context.Countries on userCountryRole.Country_Id equals country.Id
                      where country.Id == countryId && businessConfig.BusinessDay == dayOfTheMonth
                      select new BusinessConfiguration()
                      {
                          RuleId = businessConfig.Id,
                          Rule = businessConfig.Name,
                          Owner = owner.Email,
                          Recipient = recipient.Email
                      }).FirstOrDefault();
            }
            return this;
        }
        public bool ExecuteBusinessRule(BusinessConfiguration businessConfiguration)
        {

            // If the rule execution is successful

            return true;
        }

    }
}
