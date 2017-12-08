using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskUtility
{
    public class Procedure
    {
        
        private Procedure() { }
        public Procedure(string command)
        {
            this.Command = command;
        }
        public int TaskID { get; set; }

        public TaskTypes TaskName { get; set; }

        public string Command { get; set; }

        public void ExecuteTask()
        {
            Console.WriteLine(Command+":" + TaskName.ToString() + ":" + TaskID.ToString());
        }
        
    }
}
