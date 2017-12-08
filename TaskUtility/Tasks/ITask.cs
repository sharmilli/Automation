
namespace TaskUtility
{
    public interface ITask
    {
        int TaskID { get; set; }

        TaskTypes TaskName { get; set; }

        void ExecuteTask();
    }
}
