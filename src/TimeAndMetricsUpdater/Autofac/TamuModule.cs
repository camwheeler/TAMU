using System;
using System.IO;
using System.Windows.Forms;
using Autofac;
using Autofac.Extras.Quartz;
using Google.GData.Spreadsheets;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Quartz;
using Quartz.Impl;
using TimeAndMetricsUpdater.Entities;

namespace TimeAndMetricsUpdater.Autofac {
    public class TamuModule : Module {
        protected override void Load(ContainerBuilder builder){
            builder.RegisterModule(new QuartzAutofacFactoryModule());
            builder.RegisterType<ProcessIcon>().As<ITrayIcon>();
            builder.RegisterType<TimeSync>().As<ISyncTime>().SingleInstance();
            builder.RegisterType<NotifyIcon>().AsSelf();
            builder.RegisterType<AutoSubmit>().AsSelf().SingleInstance();

            builder.Register(ctx => new SpreadsheetsService("TimeAndMetrics") {RequestFactory = OAuth2.GetOAuthFactory()}).As<SpreadsheetsService>();
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var tamuData = appData + @"/TAMU";
            var tamuDataUser = tamuData + @"/user.info";
            builder.Register(ctx => File.Exists(tamuDataUser) ?
                new Data {User = JsonConvert.DeserializeObject<UserInfo>(File.ReadAllText(tamuDataUser))} :
                new Data {
                    User = new UserInfo {
                        Name = Interaction.InputBox("Enter your name as it appears in the Time and Metrics sheet", "Time and Metrics Sync: User Name"),
                        GrindstoneDB = string.Format("Data Source={0};Persist Security Info=False;", Interaction.InputBox("Enter the path to your Grindstone 3 database.", "Time and Metrics Sync: DB Path", string.Format(@"{0}\Databases\Grindstone3.gsdb", appData))),
                        SheetName = "project time & metrics 2015"
                    }
                })
                .As<IData>()
                .SingleInstance()
                .OnActivated(data =>{
                    if (!Directory.Exists(tamuData))
                        Directory.CreateDirectory(tamuData);
                    File.WriteAllText(tamuDataUser, JsonConvert.SerializeObject(data.Instance.User));
                });
        }
    }
}
