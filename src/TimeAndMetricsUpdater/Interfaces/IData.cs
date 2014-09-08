using System.Collections.Generic;
using TimeAndMetricsUpdater.Entities;

namespace TimeAndMetricsUpdater
{
    public interface IData
    {
        UserInfo User { get; set; }
        List<TaskList> Tasks { get; set; }
    }
}