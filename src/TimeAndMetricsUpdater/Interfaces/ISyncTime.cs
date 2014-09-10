using System;

namespace TimeAndMetricsUpdater
{
    public interface ISyncTime
    {
        void InsertTime();
        void InsertTime(object sender, EventArgs e);
        void UpdateCategories();
        void UpdateCategories(object sender, EventArgs e);
    }
}