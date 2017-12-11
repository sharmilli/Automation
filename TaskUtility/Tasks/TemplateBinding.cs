using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskUtility.Tasks
{
    public class TemplateBinding
    {
        private TemplateBinding() { }
        public TemplateBinding(string reportId,int countryId)
        {

        }
        public int TaskID { get; set; }

        public TaskTypes TaskName { get; set; }

        public string Command { get; set; }

        public void ExecuteTask()
        {
            
        }
    }
}
