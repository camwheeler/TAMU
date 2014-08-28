using System;
using DapperExtensions.Mapper;

namespace TimeAndMetricsUpdater.Entities
{
    public class TaskList
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public bool TaskForceCrap { get; set; }
        public uint Row { get; set; }
    }

    public class TaskListMap : ClassMapper<TaskList>
    {
        public TaskListMap() {
            Table("Tasks");
            Map(t => t.Id).Key(KeyType.Guid);
            Map(t => t.Name);
            Map(t => t.TaskForceCrap).Column("TaskForceIgnoreUnsubmitted");
        }
    }
}