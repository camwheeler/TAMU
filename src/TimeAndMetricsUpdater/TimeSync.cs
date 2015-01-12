using System;
using System.Collections.Generic;
using System.Data.SqlServerCe;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DapperExtensions;
using Google.GData.Spreadsheets;
using TimeAndMetricsUpdater.Entities;

namespace TimeAndMetricsUpdater.Autofac
{
    public class TimeSync : ISyncTime
    {
        private readonly IData data;
        private readonly SpreadsheetsService projectTime;

        public TimeSync(IData data, SpreadsheetsService projectTime){
            this.data = data;
            this.projectTime = projectTime;
        }

        public void InsertTime() { 
            using (var connection = new SqlCeConnection(data.User.GrindstoneDB)) {
                try{
                    var endTime = DateTime.Now.StartOfWeek(DayOfWeek.Sunday);
                    var startTime = endTime.AddDays(-7);
                    var result = connection.Query<TaskTime>("select TaskId, Name, Start, [End] from Tasks,Times where Times.TaskId = Tasks.Id and Start > @TwoSundaysAgo and [End] < @LastSunday", new {TwoSundaysAgo = startTime, LastSunday = endTime}).ToList();

                    var elapsed = result.GroupBy(r => r.Name).Select(g => new {Name = g.First().Name, TotalTime = new TimeSpan(g.Sum(s => s.Elapsed.Ticks))});


                    var query = new SpreadsheetQuery();
                    var feed = projectTime.Query(query);
                    var spreadsheet = (SpreadsheetEntry) feed.Entries.Single(f => f.Title.Text == data.User.SheetName);
                    var worksheet = GetPreviousSheet(spreadsheet.Worksheets);
                    var cellQuery = new CellQuery(worksheet.CellFeedLink) {
                        MaximumColumn = 12,
                        MinimumRow = 11,
                        ReturnEmpty = ReturnEmptyCells.yes
                    };
                    var cells = projectTime.Query(cellQuery).Entries.Select(c => (CellEntry) c).ToList();

                    var myCell = GetMatchingCell(data.User.Name, cells);
                    foreach (var item in elapsed){
                        var row = GetMatchingCell(item.Name, cells).Row;
                        var cell = cells.Single(c => c.Row == row && c.Column == myCell.Column);
                        cell.InputValue = item.TotalTime.TotalHours.ToString("0.0");
                        cell.Update();
                    }
                }
                catch (Exception ex) { }
                UpdateCategories();
            }
        }

        public void InsertTime(object sender, EventArgs e){
            InsertTime();
        }

        private static CellEntry GetMatchingCell(string cellContents, List<CellEntry> cells) {
            var entry = new CellEntry();
            Parallel.ForEach(cells.AsParallel(), (cell, loopState) => {
                if (cell.Value == cellContents) {
                    entry = cell;
                    loopState.Stop();
                }
            });
            return entry;
        }

        public void UpdateCategories() { 
            data.Tasks = new List<TaskList>();

            // Make the request to Google
            var query = new SpreadsheetQuery();
            var feed = projectTime.Query(query);
            var spreadsheet = (SpreadsheetEntry)feed.Entries.Single(f => f.Title.Text == data.User.SheetName);
            var worksheet = GetCurrentSheet(spreadsheet.Worksheets);
            var cellQuery = new CellQuery(worksheet.CellFeedLink) {
                MaximumColumn = 1, 
                MinimumRow = 11, 
                ReturnEmpty = ReturnEmptyCells.no
            };
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
                data.Tasks.Add(new TaskList { Name = cell.Value, Row = cell.Row, TaskForceCrap = false });
            }

            using (var connection = new SqlCeConnection(data.User.GrindstoneDB)) {
                var existingTasks = connection.Query<TaskList>("select Id, Name from Tasks") ?? new List<TaskList>();
                var deleteme = existingTasks.Where(e => !data.Tasks.Select(d => d.Name).Contains(e.Name)).ToList();
                deleteme.ForEach(d =>{
                    connection.Execute("delete Times where TaskId = @Id", new{d.Id});
                    connection.Delete(d);
                });
                connection.Insert(data.Tasks.Where(t => existingTasks.All(e => e.Name != t.Name)));
            }
        }

        public void UpdateCategories(object sender, EventArgs e){
            UpdateCategories();
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
    }
}