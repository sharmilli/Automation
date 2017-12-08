using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using TaskUtility.DataAccess;
namespace TaskUtility
{
    public class CarryForwardBusinessConfiguration : ICarryForwardBusinessConfiguration
    {
        private CarryForwardBusinessConfiguration() { }
        public CarryForwardBusinessConfiguration(int countryId) {
            ExecuteCarryForwardTasks(countryId);
        }
        public CarryForwardBusinessConfiguration(IBusinessConfiguration businessConfiguration) {
            AddCarryForwardRule(businessConfiguration);
        }
        
        public int Id { get; set; }
        public long RuleId { get; set; }

        public int CountryId { get; set; }
        public DateTime PendingSince { get; set; }
        
        public int NthWorkingDayOfMonth { get; set; }

        public int RemainingWorkingDaysOfMonth { get; set; }

        [DataType(DataType.EmailAddress)]
        public string Recipient { get; set; }
        [DataType(DataType.EmailAddress)]
        public string Owner { get; set; }

        public string Rule { get; set; }

        public Freequency RuleFreequency { get; set; }

        public void ExecuteCarryForwardTasks(int countryId)
        {
            List<CarryForwardBusinessConfiguration> carryForwardRules = new List<CarryForwardBusinessConfiguration>();
            //connect to the database and fetch the list of carry forward rules
            using (TaskUtilityAutomationContainer context = new TaskUtilityAutomationContainer())
            {
                carryForwardRules = (from carryRules in context.CarryForwardRules
                                     join rules in context.Rules on carryRules.Rule_Id equals rules.Id
                                     join owner1 in context.UserCountryRoles on rules.UserCountryRole_Id equals owner1.Id
                                     join recipient1 in context.UserCountryRoles on rules.UserCountryRole1_Id equals recipient1.Id
                                     join owner in context.Users on owner1.User_Id equals owner.Id
                                     join recipient in context.Users on recipient1.User_Id equals recipient.Id
                                     join country in context.Countries on owner1.Country_Id equals country.Id
                                     where country.Id == countryId
                                     select new CarryForwardBusinessConfiguration
                                     {
                                         RuleId = rules.Id,
                                         NthWorkingDayOfMonth = rules.BusinessDay,//Will be used while logging
                                         PendingSince = carryRules.PendingSince//will be used while logging
                                     }).ToList();
            }
            BusinessConfiguration bc;
            System.Threading.Tasks.Parallel.ForEach<CarryForwardBusinessConfiguration>(carryForwardRules, (carryRule) => {
                bc = new BusinessConfiguration(countryId, carryRule.NthWorkingDayOfMonth, false);
                if (bc.ExecuteBusinessRule(bc))
                {
                    DeleteCarryForwardRule(carryRule);
                }else
                {
                    //logging
                }
            });
            
        }
        private void AddCarryForwardRule(IBusinessConfiguration businessConfiguration)
        {
            // If the rule execution is not successful, add record the carryforward table
            using (TaskUtilityAutomationContainer context = new TaskUtilityAutomationContainer())
            {
                var carryForward = new CarryForwardRule();
                carryForward.PendingSince = DateTime.Now;
                carryForward.Rule_Id = businessConfiguration.RuleId;
                context.CarryForwardRules.Add(carryForward);
                context.SaveChanges();


            }
        }
        private void DeleteCarryForwardRule(ICarryForwardBusinessConfiguration carryFwdConfig)
        {
            // If pending rule execution is successfull, the corresponding carry forward rule will be deleted from the carryforward table.
            using (TaskUtilityAutomationContainer context = new TaskUtilityAutomationContainer())
            {

                var record = context.CarryForwardRules.Where(e => e.Id.Equals(carryFwdConfig.Id)).FirstOrDefault();
                context.CarryForwardRules.Remove(record);
                context.SaveChanges();


            }
        }
    }
}
