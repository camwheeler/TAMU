using System;
using System.Drawing;
using System.Windows.Forms;

namespace TimeAndMetricsUpdater
{
    public class ProcessIcon : ITrayIcon
    {
        private readonly ISyncTime sync;
        private readonly NotifyIcon ni;

        public ProcessIcon(NotifyIcon ni, ISyncTime sync) {
            this.sync = sync;
            this.ni = ni;
        }

        public void Display() {
            ToolStripItem toolStripUpdateCategories = new ToolStripMenuItem("Update Categories (Destructive)");
            toolStripUpdateCategories.Click += sync.UpdateCategories;

            ToolStripItem toolStripInsertTime = new ToolStripMenuItem("Insert Last Week's Time");
            toolStripInsertTime.Click += sync.InsertTime;

            ToolStripItem toolStripExit = new ToolStripMenuItem("Exit");
            toolStripExit.Click += Exit;

            var ctxMenuStrip = new ContextMenuStrip();
            ctxMenuStrip.Items.Add(toolStripUpdateCategories);
            ctxMenuStrip.Items.Add(toolStripInsertTime);
            ctxMenuStrip.Items.Add("-");
            ctxMenuStrip.Items.Add(toolStripExit);
            ni.ContextMenuStrip = ctxMenuStrip;
            ni.Text = "Time And Metrics Updater";
            ni.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "alarm.ico", 24, 24);
            ni.Visible = true;
        }

        public void Dispose() {
            ni.Dispose();
        }

        private static void Exit(object sender, EventArgs e) {
            Application.Exit();
        }
    }
}
