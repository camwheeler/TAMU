using System;
using System.Windows.Forms;

namespace TimeAndMetricsUpdater
{
    class Program
    {
        [STAThread]
        static void Main(string[] args) {
            using (var pi = new ProcessIcon()) {
                pi.Display();
                Application.Run();
            }
        }
    }
}
