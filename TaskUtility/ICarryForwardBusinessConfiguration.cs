using System;

namespace TaskUtility
{
    public interface ICarryForwardBusinessConfiguration : IBusinessConfiguration
    {
        int Id { get; set; }
        DateTime PendingSince { get; set; }
    }
}
