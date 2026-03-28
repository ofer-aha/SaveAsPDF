using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace OutlookLauncher
{
    class Program
    {
        static void Main(string[] args)
        {
            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        }
    }
}