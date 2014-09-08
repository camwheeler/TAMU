using System.Collections.Generic;
using TimeAndMetricsUpdater.Entities;

namespace TimeAndMetricsUpdater.Autofac
{
    public class Data
    {
        public UserInfo User { get; set; }
        public List<TaskList> Tasks { get; set; }
    }
}