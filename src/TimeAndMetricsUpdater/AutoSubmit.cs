using System;
using System.Windows.Forms;
using Quartz;

namespace TimeAndMetricsUpdater
{
    public class AutoSubmit : IJob {
        private readonly ISyncTime sync;

        public AutoSubmit(ISyncTime sync) {
            this.sync = sync;
            sync.UpdateCategories();
        }

        public void Execute(IJobExecutionContext context){
            //sync.InsertTime();
            MessageBox.Show("Job triggered.", "TAMU", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}