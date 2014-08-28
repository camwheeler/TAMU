using System;

namespace TimeAndMetricsUpdater.Entities
{
    public class TaskTime
    {
        public Guid TaskId { get; set; }
        public string Name { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public TimeSpan Elapsed {get{return End.Subtract(Start);}}
    }
}