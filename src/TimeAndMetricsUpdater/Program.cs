using System;
using System.Windows.Forms;
using Autofac;
using Quartz;
using TimeAndMetricsUpdater.Autofac;

namespace TimeAndMetricsUpdater
{
    public class Program
    {
        private static void Main(string[] args){
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            var builder = new ContainerBuilder();
            builder.RegisterModule<TamuModule>();
            var container = builder.Build();
            using (var scope = container.BeginLifetimeScope()){

                var job = JobBuilder.Create<AutoSubmit>()
                    .WithIdentity("AutoInsert", "Weekly")
                    .Build();

                var trigger = TriggerBuilder.Create()
                    .WithIdentity("InsertTime", "Weekly")
                    .WithSchedule(CronScheduleBuilder.WeeklyOnDayAndHourAndMinute(DayOfWeek.Monday, 0, 1))
                    .Build();

                var scheduler = scope.Resolve<IScheduler>();
                scheduler.ScheduleJob(job, trigger);
                scheduler.Start();

                using (var pi = scope.Resolve<ITrayIcon>()){
                    pi.Display();

                    Application.Run();
                }
                scheduler.Shutdown();
            }
        }
    }
}
