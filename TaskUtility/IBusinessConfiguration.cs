namespace TaskUtility
{
    public interface IBusinessConfiguration
    {
        long RuleId { get; set; }
        int CountryId { get; set; }
        string Rule { get; set; }
        string Owner { get; set; }
        int NthWorkingDayOfMonth { get; set; }
        int RemainingWorkingDaysOfMonth { get; set; }
        string Recipient { get; set; }
        Freequency RuleFreequency { get; set; }
    }
}
