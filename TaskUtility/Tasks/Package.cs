
namespace TaskUtility
{
    public class Package
    {
        private Package() { }

        public Package(string name,string command)
        {
            Name = name;
            Command = command;
        }
        public int TaskID { get; set; }

        public TaskTypes TaskName { get; set; }

        public string Name { get; set; }

        public string Command { get; set; }

        public void ExecuteTask()
        {
             //The logic to execute the desired SSIS package goes here. If the action is successful return
        }
    }
}
