using System;
using System.Security.Cryptography.X509Certificates;
using Google.GData.Spreadsheets;

namespace TimeAndMetricsUpdater
{
    public interface ITrayIcon : IDisposable
    {
        void Display();
    }
}