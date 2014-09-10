using System;
using System.Windows.Forms;
using Autofac;
using Quartz;
using Quartz.Spi;
using TimeAndMetricsUpdater.Autofac;

namespace TimeAndMetricsUpdater
{
    public class Program
    {
        static void Main(string[] args) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var builder = new ContainerBuilder();
            builder.RegisterModule<TamuModule>();
            var container = builder.Build();
            var scheduler = container.Resolve<IScheduler>();
            scheduler.Start();
            
            var job = JobBuilder.Create<AutoSubmit>()
                .WithIdentity("AutoInsert", "Weekly")
                .Build();

            var trigger = TriggerBuilder.Create()
                .WithIdentity("InsertTime", "Weekly")
                .StartNow()
                .WithSchedule(CronScheduleBuilder.WeeklyOnDayAndHourAndMinute(DayOfWeek.Monday, 0, 1))
                .Build();

            scheduler.ScheduleJob(job, trigger);

            var asfd = trigger.GetNextFireTimeUtc();

            using (var pi = container.Resolve<ITrayIcon>()){
                pi.Display();

                Application.Run();
            }
        }
    }
}
