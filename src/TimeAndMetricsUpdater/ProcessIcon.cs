using System;
using System.Collections.Generic;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dapper;
using DapperExtensions;
using Google.GData.Spreadsheets;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using TimeAndMetricsUpdater.Entities;

namespace TimeAndMetricsUpdater
{
    public class ProcessIcon : IDisposable
    {
        private readonly IData data;
        private readonly NotifyIcon ni;

        public ProcessIcon(NotifyIcon ni, IData data) {
            this.data = data;
            this.ni = ni;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }

        public void Dispose() {
            ni.Dispose();
        }

        public void Display() {
            ToolStripItem toolStripUpdateCategories = new ToolStripMenuItem("Update Categories");
            toolStripUpdateCategories.Click += UpdateCategories;

            ToolStripItem toolStripInsertTime = new ToolStripMenuItem("Insert Last Week's Time");
            toolStripInsertTime.Click += InsertTime;

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

            if (File.Exists("user.info"))
                data.User = JsonConvert.DeserializeObject<UserInfo>(File.ReadAllText("user.info"));
            else{
                var defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\Databases\Grindstone3.gsdb";
                data.User = new UserInfo {
                    Name = Interaction.InputBox("Enter your name as it appears in the Time and Metrics sheet", "Time and Metrics Sync: User Name"), 
                    GrindstoneDB = "Data Source=" + Interaction.InputBox("Enter the path to your Grindstone 3 database.", "Time and Metrics Sync: DB Path", defaultPath) + @";Persist Security Info=False;"
                };
                File.WriteAllText("user.info", JsonConvert.SerializeObject(userInfo));
            }
        }

        

        private static void Exit(object sender, EventArgs e) {
            Application.Exit();
        }
    }
}