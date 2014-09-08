using System;
using System.IO;
using Autofac;
using Google.GData.Spreadsheets;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using TimeAndMetricsUpdater.Entities;

namespace TimeAndMetricsUpdater.Autofac {
    public class TamuModule : Module {
        protected override void Load(ContainerBuilder builder){
            builder.RegisterType<ProcessIcon>().As<ITrayIcon>();
            builder.RegisterType<TimeSync>().As<ISyncTime>();

            builder.Register(ctx => new SpreadsheetsService("TimeAndMetrics") {RequestFactory = OAuth2.GetOAuthFactory()}).As<SpreadsheetsService>();
            builder.Register(ctx =>{
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                return File.Exists(appData + @"TAMU\user.info") ?
                    new Data {User = JsonConvert.DeserializeObject<UserInfo>(File.ReadAllText("user.info"))}
                    : new Data {
                        User = new UserInfo {
                            Name = Interaction.InputBox("Enter your name as it appears in the Time and Metrics sheet", "Time and Metrics Sync: User Name"),
                            GrindstoneDB = string.Format("Data Source={0};Persist Security Info=False;", Interaction.InputBox("Enter the path to your Grindstone 3 database.", "Time and Metrics Sync: DB Path", string.Format(@"{0}\Databases\Grindstone3.gsdb", appData)))
                        }
                    };
            })
                .As<IData>()
                .SingleInstance()
                .OnActivated(data => File.WriteAllText("user.info", JsonConvert.SerializeObject(data.Instance.User)));
        }
    }
}
