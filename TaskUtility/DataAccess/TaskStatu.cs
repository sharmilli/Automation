//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TaskUtility.DataAccess
{
    using System;
    using System.Collections.Generic;
    
    public partial class TaskStatu
    {
        public long Id { get; set; }
        public int Task_Id { get; set; }
        public long Status_Id { get; set; }
    
        public virtual Status Status { get; set; }
        public virtual Task Task { get; set; }
    }
}
