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

        private static void InsertTime(object sender, EventArgs eventArgs) {
            using (var connection = new SqlCeConnection(userInfo.GrindstoneDB)) {
                var endTime = DateTime.Now.StartOfWeek(DayOfWeek.Sunday);
                var startTime = endTime.AddDays(-7);
                var result = connection.Query<TaskTime>("select TaskId, Name, Start, [End] from Tasks,Times where Times.TaskId = Tasks.Id and Start > @TwoSundaysAgo and [End] < @LastSunday", new { TwoSundaysAgo = startTime, LastSunday = endTime }).ToList();

                var elapsed = result.GroupBy(r => r.Name).Select(g => new{ Name = g.First().Name, TotalTime = new TimeSpan(g.Sum(s => s.Elapsed.Ticks))});


                var query = new SpreadsheetQuery();
                var feed = projectTime.Query(query);
                var spreadsheet = (SpreadsheetEntry)feed.Entries.Single(f => f.Title.Text == "project time & metrics 2014");
                var worksheet = GetPreviousSheet(spreadsheet.Worksheets);
                var cellQuery = new CellQuery(worksheet.CellFeedLink) {
                    MaximumColumn = 12, 
                    MinimumRow = 11, 
                    ReturnEmpty = ReturnEmptyCells.yes
                };
                var cells = projectTime.Query(cellQuery).Entries.Select(c => (CellEntry)c).ToList();

                var myCell = GetMatchingCell(userInfo.Name, cells);
                foreach (var item in elapsed){
                    var row = GetMatchingCell(item.Name, cells).Row;
                    var cell = cells.Single(c => c.Row == row && c.Column == myCell.Column);
                    cell.InputValue = item.TotalTime.TotalHours.ToString("0.0");
                    cell.Update();
                }
                
            }
        }

        private static CellEntry GetMatchingCell(string cellContents, List<CellEntry> cells){
            var entry = new CellEntry();
            Parallel.ForEach(cells.AsParallel(), (cell, loopState) =>{
                if (cell.Value == cellContents) {
                    entry = cell;
                    loopState.Stop();
                }
            });
            return entry;
        }

        private static void UpdateCategories(object sender, EventArgs eventArgs) {
            tasks = new List<TaskList>();

            // Make the request to Google
            // See other portions of this guide for code to put here...

            var query = new SpreadsheetQuery();
            var feed = projectTime.Query(query);
            var spreadsheet = (SpreadsheetEntry)feed.Entries.Single(f => f.Title.Text == "project time & metrics 2014");
            var worksheet = GetCurrentSheet(spreadsheet.Worksheets);
            var cellQuery = new CellQuery(worksheet.CellFeedLink);
            cellQuery.MaximumColumn = 1;
            cellQuery.MinimumRow = 11;
            cellQuery.ReturnEmpty = ReturnEmptyCells.no;
            var cells = projectTime.Query(cellQuery).Entries.Select(c => (CellEntry)c).ToList();

            var categories = new List<string>() { 
                "Administration/Support", 
                "Operations Projects", 
                "Customer Projects", 
                "Strategic Projects", 
                "Software Upgrades", 
                "TOTALS", 
                "INTERNAL OPS/ADMINISTRATION", 
                "OPERATIONS PROJECTS", 
                "CUSTOMER PROJECTS", 
                "STRATEGIC PROJECTS", 
                "SOFTWARE UPGRADES", 
                "CARD@ONCE", 
                "EMV", 
                "CLIENT IMPLEMENTATIONS",
                "AUDIT REMEDIATION"
            };

            foreach (var cell in cells.Where(c => !categories.Contains(c.Value))) {
                tasks.Add(new TaskList { Name = cell.Value, Row = cell.Row, TaskForceCrap = false });
            }

            using (var connection = new SqlCeConnection(userInfo.GrindstoneDB)) {
                var existingTasks = connection.Query<TaskList>("select Id, Name from Tasks").ToList();
                connection.Insert(tasks.Where(t => existingTasks.All(e => e.Name != t.Name)));
            }
        }

        private static WorksheetEntry GetCurrentSheet(WorksheetFeed feed) {
            var currentMonth = DateTime.Now.Month;
            var currentDay = DateTime.Now.Day;

            foreach (var entry in feed.Entries) {
                if (entry.Title.Text.Contains(".")) {
                    var month = Int32.Parse(entry.Title.Text.Substring(0, entry.Title.Text.IndexOf(".")));
                    var day = Int32.Parse(entry.Title.Text.Substring(entry.Title.Text.IndexOf(".") + 1));
                    if (month == currentMonth)
                        if (day >= currentDay)
                            return (WorksheetEntry)entry;
                    if (month > currentMonth)
                        return (WorksheetEntry)entry;
                }
            }
            throw new Exception("Can't find a current page!");
        }

        private static WorksheetEntry GetPreviousSheet(WorksheetFeed feed) {
            var previousSaturday = DateTime.Now.AddDays(-7);
            while (previousSaturday.DayOfWeek != DayOfWeek.Saturday)
            {
                previousSaturday = previousSaturday.AddDays(1);
            }

            var previousEntry = feed.Entries.SingleOrDefault(e => e.Title.Text == string.Format("{0}.{1}", previousSaturday.Month, previousSaturday.Day));
            if(previousEntry == null)
                throw new Exception("Can't find the previous page!");

            return (WorksheetEntry)previousEntry;
        }

        private static void Exit(object sender, EventArgs e) {
            Application.Exit();
        }
    }
}
