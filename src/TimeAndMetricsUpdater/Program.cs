using System;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using Autofac;

namespace TimeAndMetricsUpdater
{
    public class Program
    {
        private static IContainer Container { get; set; }

        [STAThread]
        static void Main(string[] args) {
            using (var pi = new ProcessIcon()) {
                pi.Display();

                var autoSubmit = new AutoSubmit();
                autoSubmit.Start();

                Application.Run();
            }
        }
    }
}
